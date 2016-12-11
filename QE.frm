VERSION 5.00
Begin VB.Form frmStandard 
   BackColor       =   &H00000000&
   Caption         =   "OKOLI, Lotanna Uche - Calculator Project - Standard Mode"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboMode 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdMemoryPlus 
      Caption         =   "&M+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   34
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdMemoryMinus 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   33
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdMemoryRecall 
      Caption         =   "M&R"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   32
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdPie 
      Caption         =   "&Pi"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   31
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdSine 
      Caption         =   "&Sin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   30
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdCosine 
      Caption         =   "&Cos"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   29
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdTangent 
      Caption         =   "&Tan"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   28
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdLogBase10 
      Caption         =   "&Log"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   27
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdSquare 
      Caption         =   "&x^2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   26
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdReciprocal 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   25
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdSqrt 
      Caption         =   "S&qrt"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   24
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdAns 
      Caption         =   "A&ns"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdClearAll 
      BackColor       =   &H000080FF&
      Caption         =   "&AC"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      MaskColor       =   &H000080FF&
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Del"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   17
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtPreAns 
      Height          =   735
      Left            =   6600
      TabIndex        =   16
      Text            =   "0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdAnswer 
      Caption         =   "&="
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   14
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "&-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "&*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "&/"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdZero 
      Caption         =   "&0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "&."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "&2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdOne 
      Caption         =   "&1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdThree 
      Caption         =   "&3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "&5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdFour 
      Caption         =   "&4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdSix 
      Caption         =   "&6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton cmdEight 
      Caption         =   "&8"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdSeven 
      Caption         =   "&7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNine 
      Caption         =   "&9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox txtOperator 
      Height          =   375
      Left            =   8040
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAns 
      Height          =   615
      Left            =   7920
      TabIndex        =   19
      Text            =   "0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtMemory 
      Height          =   375
      Left            =   8280
      TabIndex        =   35
      Text            =   "0"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblMode 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   10200
      TabIndex        =   37
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblMemory 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAnswer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   22
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label lblScreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   21
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmStandard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Answers()
Dim Ans As Double
If txtOperator = "+" Then
Ans = CDbl(txtPreAns.Text) + CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator = "*" Then
Ans = CDbl(txtPreAns.Text) * CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator = "-" Then
Ans = CDbl(txtPreAns.Text) - CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator = "/" Then
Ans = CDbl(txtPreAns.Text) / CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
End If
End If
End If
End If
End Function

Private Sub cboMode_Click()
If cboMode.ListIndex = 1 Then
frmQE.Show
frmStandard.Hide
frmSE.Hide
frmQE.cboMode.ListIndex = 1
frmQE.lblX1.Caption = ""
frmQE.lblX2.Caption = ""
frmQE.txtA.Text = ""
frmQE.txtB.Text = ""
frmQE.txtC.Text = ""
Else
If cboMode.ListIndex = 0 Then
frmQE.Hide
frmStandard.Show
frmSE.Hide
frmStandard.cboMode.ListIndex = 0
frmStandard.lblScreen.Caption = ""
frmStandard.lblAnswer.Caption = ""
Else
If cboMode.ListIndex = 2 Then
frmQE.Hide
frmStandard.Hide
frmSE.Show
frmSE.cboMode.ListIndex = 2
frmSE.txtA1.Text = ""
frmSE.txtA2.Text = ""
frmSE.txtB1.Text = ""
frmSE.txtb2.Text = ""
frmSE.txtC1.Text = ""
frmSE.txtC2.Text = ""
frmSE.lblX.Caption = ""
frmSE.lblY.Caption = ""
End If
End If
End If
End Sub

Private Sub CmdAdd_Click()
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
End If
End If
lblScreen.Caption = lblScreen.Caption & "+"
txtOperator.Text = "+"
End Sub

Private Sub cmdAns_Click()
If Right(lblScreen.Caption, 1) = "=" Then
lblAnswer.Caption = CDbl(txtAns.Text)
lblScreen.Caption = "Ans"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = CDbl(txtAns.Text)
lblScreen.Caption = lblScreen.Caption & "Ans"
Else
lblScreen.Caption = "Ans"
lblAnswer.Caption = CDbl(txtAns.Text)
End If
End If
End Sub

Private Sub cmdAnswer_Click()
txtAns.Text = CDbl(lblAnswer.Caption)
If Right(lblScreen.Caption, 1) = "=" Then
txtAns.Text = CDbl(lblAnswer.Caption)
lblAnswer.Caption = txtAns.Text
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Dim Ans As Double
If txtOperator.Text = "+" Then
Ans = CDbl(txtPreAns.Text) + CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator.Text = "-" Then
Ans = CDbl(txtPreAns.Text) - CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator.Text = "*" Then
Ans = CDbl(txtPreAns.Text) * CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
If txtOperator.Text = "/" Then
Ans = CDbl(txtPreAns.Text) / CDbl(lblAnswer.Caption)
txtAns.Text = Ans
lblAnswer.Caption = txtAns.Text
Else
txtAns.Text = CDbl(lblAnswer.Caption)
lblAnswer.Caption = txtAns.Text
End If
End If
End If
End If
End If
End If
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = lblScreen.Caption & ""
Else
Dim OneKey As String
OneKey = lblScreen.Caption
lblScreen.Caption = OneKey & "="
End If
End Sub

Private Sub cmdClear_Click()
If lblAnswer.Caption = "" Then
lblAnswer.Caption = lblAnswer.Caption & ""
Else
lblAnswer.Caption = Left(lblAnswer.Caption, Len(lblAnswer.Caption) - 1)
End If
If lblScreen.Caption = "" Then
lblScreen.Caption = lblScreen.Caption & ""
Else
lblScreen.Caption = Left(lblScreen.Caption, Len(lblScreen.Caption) - 1)
End If
End Sub

Private Sub cmdClearAll_Click()
lblScreen.Caption = ""
lblAnswer.Caption = ""
txtPreAns.Text = "0"
txtOperator.Text = ""
End Sub

Private Sub cmdCosine_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "cos(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Cos(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
Else
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "cos(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Cos(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End Sub

Private Sub cmdDecimal_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "."
lblScreen.Caption = "."
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "."
lblScreen.Caption = lblScreen.Caption & "."
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "."
lblScreen.Caption = lblScreen.Caption & "."
Else
lblAnswer.Caption = lblAnswer.Caption & "."
lblScreen.Caption = lblScreen.Caption & "."
End If
End If
End If
End Sub

Private Sub cmdDivide_Click()
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
End If
End If
lblScreen.Caption = lblScreen.Caption & "/"
txtOperator.Text = "/"
End Sub

Private Sub cmdEight_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "8"
lblScreen.Caption = "8"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "8"
lblScreen.Caption = lblScreen.Caption & "8"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "8"
lblScreen.Caption = lblScreen.Caption & "8"
Else
lblAnswer.Caption = lblAnswer.Caption & "8"
lblScreen.Caption = lblScreen.Caption & "8"
End If
End If
End If
End Sub

Private Sub cmdFive_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "5"
lblScreen.Caption = "5"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "5"
lblScreen.Caption = lblScreen.Caption & "5"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "5"
lblScreen.Caption = lblScreen.Caption & "5"
Else
lblAnswer.Caption = lblAnswer.Caption & "5"
lblScreen.Caption = lblScreen.Caption & "5"
End If
End If
End If
End Sub

Private Sub cmdFour_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "4"
lblScreen.Caption = "4"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "4"
lblScreen.Caption = lblScreen.Caption & "4"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "4"
lblScreen.Caption = lblScreen.Caption & "4"
Else
lblAnswer.Caption = lblAnswer.Caption & "4"
lblScreen.Caption = lblScreen.Caption & "4"
End If
End If
End If
End Sub

Private Sub cmdLogBase10_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
If Left(lblAnswer.Caption, 1) = "-" Then
MsgBox ("Cannot find the logarithm of a negative number")
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "log(" & lblAnswer.Caption & ")="
Multiply = Log(CDbl(lblAnswer.Caption)) / Log(10)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
Else
Call Answers
If Left(lblAnswer.Caption, 1) = "-" Then
lblScreen.Caption = "Ans"
MsgBox ("Cannot find the logarithm of a negative number")
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "log(" & lblAnswer.Caption & ")="
Multiply = Log(CDbl(lblAnswer.Caption)) / Log(10)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End If
End Sub

Private Sub cmdMemoryMinus_Click()
Dim PreAns As Double
Dim Memory As Double
If Right(lblScreen.Caption, 1) = "=" Then
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = Memory - PreAns
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = Memory - PreAns
Else
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = Memory - PreAns
End If
End If
If txtMemory.Text = "0" Then
lblMemory.Visible = False
Else
lblMemory.Visible = True
End If
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = lblScreen.Caption & ""
Else
lblScreen.Caption = lblScreen.Caption & "="
End If
End Sub

Private Sub cmdMemoryPlus_Click()
Dim PreAns As Double
Dim Memory As Double
If Right(lblScreen.Caption, 1) = "=" Then
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = PreAns + Memory
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = PreAns + Memory
Else
PreAns = CDbl(lblAnswer.Caption)
Memory = CDbl(txtMemory.Text)
txtMemory.Text = PreAns + Memory
End If
End If
If txtMemory.Text = "0" Then
lblMemory.Visible = False
Else
lblMemory.Visible = True
End If
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = lblScreen.Caption & ""
Else
lblScreen.Caption = lblScreen.Caption & "="
End If
End Sub

Private Sub cmdMemoryRecall_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Then
lblAnswer.Caption = CDbl(txtMemory.Text)
lblScreen.Caption = "Memory"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = CDbl(txtMemory.Text)
lblScreen.Caption = lblScreen.Caption & "Memory"
Else
OneKey = lblScreen.Caption
lblScreen.Caption = OneKey & ""
lblAnswer.Caption = lblAnswer.Caption & ""
End If
End If
txtAns.Text = lblAnswer.Caption
End Sub

Private Sub cmdMinus_Click()
Dim PreAns As Double
If lblScreen.Caption = "" And lblAnswer.Caption = "" Then
lblAnswer.Caption = "-"
Else
If Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "-"
Else
If Left(lblScreen.Caption, 1) = "-" And Len(lblScreen.Caption) < 2 Then
lblAnswer.Caption = "-"
Else
If Left(lblScreen.Caption, 2) = "-" And Len(lblScreen.Caption) < 2 Then
lblAnswer.Caption = ""
Else
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") > 0 Or InStr(1, lblScreen.Caption, "*") > 0 Or InStr(1, lblScreen.Caption, "/") > 0 Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
End If
End If
End If
End If
End If
End If
lblScreen.Caption = lblScreen.Caption & "-"
If Left(lblAnswer.Caption, 1) = "-" Then
lblAnswer.Caption = lblAnswer.Caption
Else
txtOperator.Text = "-"
End If
End Sub

Private Sub cmdMultiply_Click()
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
If InStr(1, lblScreen.Caption, "+") > 0 Or InStr(1, lblScreen.Caption, "-") Or InStr(1, lblScreen.Caption, "*") Or InStr(1, lblScreen.Caption, "/") Then
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
End If
End If
lblScreen.Caption = lblScreen.Caption & "*"
txtOperator.Text = "*"
End Sub

Private Sub cmdNine_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "9"
lblScreen.Caption = "9"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "9"
lblScreen.Caption = lblScreen.Caption & "9"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "9"
lblScreen.Caption = lblScreen.Caption & "9"
Else
lblAnswer.Caption = lblAnswer.Caption & "9"
lblScreen.Caption = lblScreen.Caption & "9"
End If
End If
End If
End Sub

Private Sub cmdOne_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "1"
lblScreen.Caption = "1"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "1"
lblScreen.Caption = lblScreen.Caption & "1"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "1"
lblScreen.Caption = lblScreen.Caption & "1"
Else
lblAnswer.Caption = lblAnswer.Caption & "1"
lblScreen.Caption = lblScreen.Caption & "1"
End If
End If
End If
End Sub

Private Sub cmdPie_Click()
If Right(lblScreen.Caption, 1) = "=" Then
lblAnswer.Caption = CDbl(4 * Atn(1))
lblScreen.Caption = "Pi"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = CDbl(4 * Atn(1))
lblScreen.Caption = lblScreen.Caption & "Pi"
Else
lblScreen.Caption = "Pi"
lblAnswer.Caption = CDbl(4 * Atn(1))
End If
End If
End Sub

Private Sub cmdReciprocal_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "reciprocal(" & lblAnswer.Caption & ")="
Multiply = Round(1 / (CDbl(lblAnswer.Caption)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
Else
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "reciprocal(" & lblAnswer.Caption & ")="
Multiply = Round(1 / (CDbl(lblAnswer.Caption)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End Sub

Private Sub cmdSeven_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "7"
lblScreen.Caption = "7"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "7"
lblScreen.Caption = lblScreen.Caption & "7"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "7"
lblScreen.Caption = lblScreen.Caption & "7"
Else
lblAnswer.Caption = lblAnswer.Caption & "7"
lblScreen.Caption = lblScreen.Caption & "7"
End If
End If
End If
End Sub

Private Sub cmdSine_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "sin(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Sin(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
Else
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "sin(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Sin(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End Sub

Private Sub cmdSix_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "6"
lblScreen.Caption = "6"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "6"
lblScreen.Caption = lblScreen.Caption & "6"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "6"
lblScreen.Caption = lblScreen.Caption & "6"
Else
lblAnswer.Caption = lblAnswer.Caption & "6"
lblScreen.Caption = lblScreen.Caption & "6"
End If
End If
End If
End Sub

Private Sub cmdSqrt_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
If Left(lblAnswer.Caption, 1) = "-" Then
MsgBox ("Cannot find the square root of negative numbers")
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "sqrt(" & lblAnswer.Caption & ")="
Multiply = Sqr(CDbl(lblAnswer.Caption))
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
Else
Call Answers
If Left(lblAnswer.Caption, 1) = "-" Then
lblScreen.Caption = "Ans"
MsgBox ("Cannot find the square root of negative numbers")
Else
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "sqrt(" & lblAnswer.Caption & ")="
Multiply = Sqr(CDbl(lblAnswer.Caption))
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End If
End Sub

Private Sub cmdSquare_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = lblAnswer.Caption & "^2="
Multiply = lblAnswer.Caption ^ 2
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
Else
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = lblAnswer.Caption & "^2="
Multiply = lblAnswer.Caption ^ 2
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End Sub

Private Sub cmdTangent_Click()
Dim Multiply As Double
Dim PreAns As Double
If Right(lblScreen.Caption, 1) = "=" Then
lblScreen.Caption = "Ans"
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "tan(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Tan(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
Else
Call Answers
PreAns = CDbl(lblAnswer.Caption)
txtPreAns.Text = PreAns
lblScreen.Caption = "tan(" & lblAnswer.Caption & ")="
Multiply = Round(CDbl(Tan(lblAnswer.Caption * 4 * Atn(1) / 180)), 15)
txtAns.Text = Multiply
lblAnswer.Caption = txtAns.Text
End If
End Sub

Private Sub cmdThree_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "3"
lblScreen.Caption = "3"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "3"
lblScreen.Caption = lblScreen.Caption & "3"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "3"
lblScreen.Caption = lblScreen.Caption & "3"
Else
lblAnswer.Caption = lblAnswer.Caption & "3"
lblScreen.Caption = lblScreen.Caption & "3"
End If
End If
End If
End Sub

Private Sub cmdTwo_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "2"
lblScreen.Caption = "2"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "2"
lblScreen.Caption = lblScreen.Caption & "2"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "2"
lblScreen.Caption = lblScreen.Caption & "2"
Else
lblAnswer.Caption = lblAnswer.Caption & "2"
lblScreen.Caption = lblScreen.Caption & "2"
End If
End If
End If
End Sub

Private Sub cmdZero_Click()
Dim OneKey As String
If Right(lblScreen.Caption, 1) = "=" Or Right(lblScreen.Caption, 1) = "s" Or Right(lblScreen.Caption, 1) = "y" Then
lblAnswer.Caption = "0"
lblScreen.Caption = "0"
Else
If Left(lblAnswer.Caption, 1) = "-" And Len(lblAnswer.Caption) < 2 Then
lblAnswer.Caption = lblAnswer.Caption & "0"
lblScreen.Caption = lblScreen.Caption & "0"
Else
If Right(lblScreen.Caption, 1) = "+" Or Right(lblScreen.Caption, 1) = "-" Or Right(lblScreen.Caption, 1) = "/" Or Right(lblScreen.Caption, 1) = "*" Then
lblAnswer.Caption = "0"
lblScreen.Caption = lblScreen.Caption & "0"
Else
lblAnswer.Caption = lblAnswer.Caption & "0"
lblScreen.Caption = lblScreen.Caption & "0"
End If
End If
End If
End Sub

Private Sub Form_Load()
cboMode.AddItem "1 - Standard"
cboMode.ItemData(cboMode.NewIndex) = 0
cboMode.AddItem "2 - Quadratic Equation"
cboMode.ItemData(cboMode.NewIndex) = 1
cboMode.AddItem "3 - Simultaneous Equation"
cboMode.ItemData(cboMode.NewIndex) = 2
cboMode.ListIndex = 0
End Sub
