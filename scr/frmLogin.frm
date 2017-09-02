VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmLogin 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Info"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Play Now"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "World War Two History Test"
      ForeColor       =   &H8000000C&
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   7215
   End
   Begin WMPLibCtl.WindowsMediaPlayer MedPlay 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12515
      _cy             =   7011
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmMain.Show
    Me.Hide
    Call LoadQuestion(1)

End Sub

Private Sub Command2_Click()

frmReading.Visible = True

End Sub

Private Sub Form_Load()

Random = CInt(Int((3 - 1 + 1) * Rnd() + 1))
MedPlay.URL = App.Path & "\Video\Intro.mp4"
MedPlay.settings.volume = 15

End Sub
