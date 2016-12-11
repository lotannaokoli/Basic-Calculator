VERSION 5.00
Begin VB.Form frmQE 
   BackColor       =   &H00000000&
   Caption         =   "OKOLI, Lotanna Uche - Calculator Project - Quadratic Equation"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
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
      Left            =   6000
      TabIndex        =   12
      Top             =   4680
      Width           =   735
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
   Begin VB.TextBox txtC 
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
      Width           =   3135
   End
   Begin VB.TextBox txtB 
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
      Width           =   3135
   End
   Begin VB.TextBox txtA 
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
      Width           =   3135
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
      Left            =   7080
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
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
   Begin VB.Label lblc 
      BackColor       =   &H80000007&
      Caption         =   "c ="
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
   Begin VB.Label lblb 
      BackColor       =   &H80000007&
      Caption         =   "b ="
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
   Begin VB.Label lblA 
      BackColor       =   &H80000007&
      Caption         =   "a ="
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
   Begin VB.Label lblX2 
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
   Begin VB.Label lblX1 
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
      Caption         =   "aX2 + bX + c = 0"
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
Attribute VB_Name = "frmQE"
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
lblX1.Caption = ""
lblX2.Caption = ""
txtA.Text = ""
txtB.Text = ""
txtC.Text = ""
End Sub

Private Sub cmdSolve_Click()
Dim A, B, C, D, E, F, G, H, X1, X2, UpOne, UpTwo, I As Double
A = CDbl(txtA.Text)
B = CDbl(txtB.Text)
C = CDbl(txtC.Text)
D = B ^ 2
E = 4 * A * C
F = D - E
G = Sqr(F)
H = 2 * A
I = -B
UpOne = I + G
UpTwo = I - G
X1 = Round((UpOne / H), 6)
X2 = Round((UpTwo / H), 6)
lblX1.Caption = "x1=" & X1
lblX2.Caption = "x2=" & X2
End Sub

Private Sub Form_Load()
cboMode.AddItem "1 - Standard"
cboMode.ItemData(cboMode.NewIndex) = 0
cboMode.AddItem "2 - Quadratic Equation"
cboMode.ItemData(cboMode.NewIndex) = 1
cboMode.AddItem "3 - Simultaneous Equation"
cboMode.ItemData(cboMode.NewIndex) = 2
cboMode.ListIndex = 1
End Sub
