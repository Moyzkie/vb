VERSION 5.00
Begin VB.Form frmCalculate 
   Caption         =   "Calculate"
   ClientHeight    =   5640
   ClientLeft      =   7275
   ClientTop       =   4710
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   5190
   Begin VB.TextBox txtQuiz2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtQuiz1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalculateAVG2 
      Caption         =   "CalculateAVG-2"
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdCalculateAVG1 
      Caption         =   "CalculateAVG-1"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label lblMessage 
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblAverage 
      Caption         =   "Average"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblQuiz2 
      Caption         =   "Quiz #2"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblQuiz1 
      Caption         =   "Quiz #1 "
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmCalculate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculateAVG1_Click()
  Dim Avg As Double
 
     MsgBox "Plese Enter The Number First!", vbInformation, "Compute Field"
     Avg = (CInt(txtQuiz1.Text) + CInt(txtQuiz2.Text)) / 2
     lblMessage.Caption = Avg
 
End Sub

Private Sub cmdCalculateAVG2_Click()
 
     MsgBox "Plese Enter The Number First!", vbInformation, "Compute Field"
     Avg = (CInt(txtQuiz1.Text) + CInt(txtQuiz2.Text)) / 2
     lblMessage.Caption = Avg

End Sub




Private Sub Form_Load()

End Sub
