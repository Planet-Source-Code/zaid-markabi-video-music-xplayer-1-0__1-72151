VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xplayer 1.0"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      DrawWidth       =   10
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2280
      ScaleHeight     =   105
      ScaleWidth      =   4185
      TabIndex        =   9
      Top             =   3000
      Width           =   4215
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   45
         TabIndex        =   10
         Top             =   0
         Width           =   50
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00C00000&
      Height          =   1785
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonOpen 
      Left            =   1800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open media file"
      FileName        =   "filename"
      Filter          =   "All Media Files"
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1200
      TabIndex        =   8
      Top             =   3300
      Width           =   735
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   3300
      Width           =   735
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Library"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   165
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   -240
      Picture         =   "frmMain.frx":57E2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   2
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Image MenuImage 
      Height          =   375
      Index           =   1
      Left            =   3720
      Picture         =   "frmMain.frx":91A6
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label Cmd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Image MenuImage 
      Height          =   375
      Index           =   0
      Left            =   2280
      Picture         =   "frmMain.frx":B046
      Top             =   3240
      Width           =   1395
   End
   Begin WMPLibCtl.WindowsMediaPlayer Xplayer 
      Height          =   2760
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4260
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7514
      _cy             =   4868
   End
   Begin VB.Image MenuImage 
      Height          =   375
      Index           =   2
      Left            =   5160
      Picture         =   "frmMain.frx":CEE6
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   -240
      Picture         =   "frmMain.frx":ED86
      Top             =   795
      Width           =   2415
   End
   Begin VB.Image MenuImage 
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "frmMain.frx":1274A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   915
   End
   Begin VB.Image MenuImage 
      Height          =   375
      Index           =   4
      Left            =   1080
      Picture         =   "frmMain.frx":145EA
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed By [ Zaid Markabi ]
' ___________________________________________________________________________________________________
'|                                                                                                   |\_______________________
'|  ###############        ###         #####   ######                ######    #####                 |                        |\0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'| ##############         #####         ###     ##   ##               ######  #####                  |      Zaid Markabi      |=\ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|         ####          ### ###        ###     ##    ##              ##  ## ##  ##                  |                        |==\0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'|       ###            ###   ###       ###     ##     ##    #####    ##   ###   ##                  | zaidmarkabi@yahoo.com  |===\ 1 __________________________________  0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 1 0 0 0 1
'|     ###             ###########      ###     ##     ##   ####      ##    #    ##                  |                        |====|>| Development For Our Digital Life | 1 1 0 0 1 1 1 0 1 0 0 1 0 0 0 1 1 0 1 0
'|   ###              #############     ###     ##    ##              ##         ##      A R K A B I | VisualBasic Programmer |===/ 1|__________________________________| 0 1 1 0 1 0 0 0 1 1 1 0 1 0 1 1 0 1 0 0
'| ##############    ###         ###    ###     ##   ##               ##         ##     ############ |                        |==/0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'| ###############   ###         ###   #####   ######                ####       ####   ### 2009 ###  |Syria(Arab Area)-Tartuse|=/ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|                                                                                    ############   | _______________________|/0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'|___________________________________________________________________________________________________|/
'|
'| FOR MORE VB APPLICATIONS ( WITH SOURCE CODES )
'| Http://www.YazanMarkabi.webs.com
'| Em@l : ZaidMarabi@yahoo.com
'|___________________________________________________________________________________________________|/


Private Sub Cmd_Click(Index As Integer)
On Error GoTo 1
Select Case Index
Case Is = 0
Xplayer.Controls.play

Case Is = 1
Xplayer.Controls.pause

Case Is = 2
Xplayer.Controls.stop

Case Is = 3
CommonOpen.ShowOpen
Xplayer.URL = CommonOpen.FileName

Case Is = 5
List1.AddItem Xplayer.URL

Case Is = 6
List1.RemoveItem List1.ListIndex


End Select

1:
End Sub

Private Sub Form_Load()
On Error Resume Next
Xplayer.URL = App.Path + "\Demo.wmv"

Dim i, num As Integer
Dim xx As String
Open App.Path + "\Library.txt" For Input As #1
Input #1, num

For i = 1 To num
Input #1, xx
List1.AddItem xx
Next

Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

Open App.Path + "\Library.txt" For Output As #1
Write #1, List1.ListCount

For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next

Close #1
End Sub

Private Sub List1_Click()
On Error Resume Next
Xplayer.URL = List1.List(List1.ListIndex)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Xplayer.Controls.currentPosition = X
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Picture1.Cls
Picture1.ScaleWidth = Xplayer.currentMedia.duration
Picture2.Left = Xplayer.Controls.currentPosition
End Sub

Private Sub Xplayer_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
Xplayer.fullScreen = True
End Sub

