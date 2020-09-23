VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A Very Simple Game"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5400
      Left            =   5520
      Picture         =   "main.frx":030A
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   5580
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   5430
      Left            =   60
      Picture         =   "main.frx":5F20C
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   60
      Width           =   5430
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   5700
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   5760
      Picture         =   "main.frx":BE10E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlag 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   5760
      Picture         =   "main.frx":BEC18
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   12
      Top             =   3420
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlag 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   5760
      Picture         =   "main.frx":BF722
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   13
      Top             =   3900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlag 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   5760
      Picture         =   "main.frx":C022C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   14
      Top             =   4380
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlagM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   6240
      Picture         =   "main.frx":C0D36
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   15
      Top             =   3420
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlagM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   6240
      Picture         =   "main.frx":C1840
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   16
      Top             =   3900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picFlagM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   6240
      Picture         =   "main.frx":C234A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   17
      Top             =   4380
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picAaa 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   5580
      Picture         =   "main.frx":C2E54
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox picAaaM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   5580
      Picture         =   "main.frx":C3936
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   19
      Top             =   5100
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox picPM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   6240
      Picture         =   "main.frx":C4418
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   11
      Top             =   2340
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picPM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   6240
      Picture         =   "main.frx":C4F22
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   10
      Top             =   1860
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picPM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   6240
      Picture         =   "main.frx":C5A2C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   9
      Top             =   1380
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picPM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   6240
      Picture         =   "main.frx":C6536
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   3
      Left            =   5760
      Picture         =   "main.frx":C7040
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   6
      Top             =   2340
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   2
      Left            =   5760
      Picture         =   "main.frx":C7B4A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   5760
      Picture         =   "main.frx":C8654
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   4
      Top             =   1380
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   5760
      Picture         =   "main.frx":C915E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   3
      Top             =   900
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   2115
      Left            =   5535
      TabIndex        =   21
      Top             =   600
      Width           =   1425
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you would like to make your own map, just edit map.bmp with a paint program :)
' - Jonas

Private Sub cmdStart_Click()
    cmdStart.Enabled = False
    Do '<-- This is what we call a "MAIN LOOP", it runs the "ticks" in a game
        DoEvents '<-- A tip: NEVER, EVER forget this in your games ;)
        
        DoKeys 'Work the key inputs
        CheckFallOff 'Check if the player fell off
        
        G1.Tag = G1.Tag + 1 'cycle the Three flag sprites
        If G1.Tag = 3 Then G1.Tag = 0
        
        PaintBoard 'Paint the Board
        
        CheckInGoal 'Check if the player is at his goal
        
        For a = 1 To 2000000 '<-- This is to make sure the simulation does not run TOO FAST
        Next a
        
        If FlagDead Then EndGame 1 'the Flag of Death is up! End the game style 1: We loose :/
        
    Loop '<-- End of the Main loop
End Sub

Private Sub Form_Load()
    ' :)
    LoadMap 'Load in the map data
    
    P1.Dire = 2 'Set the players facing direction
    
    'Print File Informaiton
    Text = Text & "Made By" & vbNewLine
    Text = Text & "Jonas Ask" & vbNewLine & vbNewLine
    Text = Text & "Absolutely no reason to vote, what so ever. This program was purely made for educational purposes." & vbNewLine
    Text = Text & ":)"
    lblInfo.Caption = Text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 End
End Sub
