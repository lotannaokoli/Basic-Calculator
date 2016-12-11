VERSION 5.00
Begin VB.Form frmSE 
   BackColor       =   &H00000000&
   Caption         =   "OKOLI, Lotanna Uche - Calculator Project - Simultaneous Equation"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTest 
      Height          =   495
      Left            =   9840
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdSolve2 
      Caption         =   "="
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
      Left            =   10080
      TabIndex        =   19
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearAll 
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
      Left            =   6120
      TabIndex        =   18
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtC2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6360
      TabIndex        =   17
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtb2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6360
      TabIndex        =   14
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtA2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6360
      TabIndex        =   13
      Top             =   2160
      Width           =   2535
   End
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
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtC1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      TabIndex        =   9
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtB1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtA1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdSolve 
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
      Left            =   7200
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblc2 
      BackColor       =   &H80000007&
      Caption         =   "c2 ="
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
      Left            =   5400
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblB2 
      BackColor       =   &H80000007&
      Caption         =   "b2 ="
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
      Left            =   5400
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblA2 
      BackColor       =   &H80000007&
      Caption         =   "a2 ="
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
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   855
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
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblc1 
      BackColor       =   &H80000007&
      Caption         =   "c1 ="
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
      Left            =   720
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblb1 
      BackColor       =   &H80000007&
      Caption         =   "b1 ="
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
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblA1 
      BackColor       =   &H80000007&
      Caption         =   "a1 ="
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
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblY 
      Alignment       =   2  'Center
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
      Left            =   4800
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
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
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblScreen 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "a1X + b1Y = C1; a2X + b2Y = C2"
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
      TabIndex        =   1
      Top             =   360
      Width           =   8175
   End
End
Attribute VB_Name = "frmSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdClearAll_Click()
txtA1.Text = ""
txtA2.Text = ""
txtB1.Text = ""
txtb2.Text = ""
txtC1.Text = ""
txtC2.Text = ""
lblX.Caption = ""
lblY.Caption = ""
End Sub

Private Sub cmdSolve_Click()
Dim a1, a2, b1, b2, c1, c2, G, E, F, H, I, J, K, L, M, X, Y As Double
a1 = CDbl(txtA1.Text)
a2 = CDbl(txtA2.Text)
b1 = CDbl(txtB1.Text)
b2 = CDbl(txtb2.Text)
c1 = CDbl(txtC1.Text)
c2 = CDbl(txtC2.Text)
E = a1 * b2
F = a2 * b1
G = E - F
H = b2 * c1
I = b1 * c2
J = H - I
X = Round((J / G), 6)
L = a1 * c2
K = a2 * c1
M = L - K
Y = Round((M / G), 6)
lblX.Caption = "x=" & X
lblY.Caption = "y=" & Y
End Sub

Private Sub cmdSolve2_Click()
Dim Up, Down, Y As Double
a1 = 1
a2 = 1
a3 = 2
b1 = 1
b2 = -2
b3 = 3
c1 = -1
c2 = 3
c3 = 1
d1 = 4
d2 = -6
d3 = 7
A = c1 * a2 * c1 * d3
B = c2 * a1 * c3 * d1
C = c1 * a3 * c2 * d1
D = c3 * a1 * c1 * d2
E = -(c1 * a2 * c3 * d1)
F = -(c2 * a1 * c1 * d3)
G = -(c1 * a3 * c1 * d2)
H = -(c3 * a1 * c2 * d1)
I = c1 * a3 * c2 * b1
J = c1 * a3 * c1 * b2
K = c2 * a1 * c3 * b1
L = c2 * a1 * c1 * b3
M = -(c3 * a1 * c2 * b1)
N = -(c3 * a1 * c1 * b2)
O = -(c1 * a2 * c3 * b1)
P = -(c1 * a2 * c1 * b3)
Up = A + B + C + D + E + F + G + H
Down = I + J + K + L + M + N + O + P
Y = Up / Down
txtTest.Text = Y
End Sub

Private Sub Form_Load()
cboMode.AddItem "1 - Standard"
cboMode.ItemData(cboMode.NewIndex) = 0
cboMode.AddItem "2 - Quadratic Equation"
cboMode.ItemData(cboMode.NewIndex) = 1
cboMode.AddItem "3 - Simultaneous Equation"
cboMode.ItemData(cboMode.NewIndex) = 2
cboMode.ListIndex = 2
End Sub
