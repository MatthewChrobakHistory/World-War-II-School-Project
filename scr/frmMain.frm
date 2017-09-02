VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Answer!"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   5640
      Width           =   2055
   End
   Begin VB.OptionButton optAns 
      Caption         =   "Option1"
      Height          =   495
      Index           =   3
      Left            =   2880
      TabIndex        =   4
      Top             =   3840
      Width           =   3135
   End
   Begin VB.OptionButton optAns 
      Caption         =   "Option1"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   3120
      Width           =   3135
   End
   Begin VB.OptionButton optAns 
      Caption         =   "Option1"
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   2400
      Width           =   3135
   End
   Begin VB.OptionButton optAns 
      Caption         =   "Option1"
      Height          =   495
      Index           =   4
      Left            =   2880
      TabIndex        =   1
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label lblQuestion 
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call modLogic.FindSelectedAnswer

End Sub

Private Sub Form_Load()

Call LoadQuestions

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub
