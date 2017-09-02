VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmEndGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Thank you for playing!"
      Height          =   735
      Left            =   1920
      TabIndex        =   2
      Top             =   6360
      Width           =   8775
   End
   Begin VB.Label Label1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   13335
   End
   Begin WMPLibCtl.WindowsMediaPlayer MedPlay 
      Height          =   6735
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   8775
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
      _cx             =   15478
      _cy             =   11880
   End
End
Attribute VB_Name = "frmEndGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

frmLogin.MedPlay.Controls.stop

MedPlay.URL = App.Path & "\Video\National Anthem.mp4"

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub
