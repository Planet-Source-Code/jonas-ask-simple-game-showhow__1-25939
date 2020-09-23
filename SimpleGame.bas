Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'these are for sound
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public Type Player
 X As Integer
 Y As Integer
 Dire As Byte
End Type
Public Type Goal
 X As Integer
 Y As Integer
 Tag As Byte
End Type


Public P1 As Player '<-- The player
Public G1 As Goal '<-- The Goal

Public Board() As Boolean
Public Width As Integer
Public Height As Integer

Public FlagDead As Boolean

Public Const Size As Byte = 30

Public Sub LoadMap()
    Main.picMap.AutoSize = False
    Main.picMap.Picture = LoadPicture(App.Path & "\map.bmp") 'load the board form file
    Main.picMap.AutoSize = True
    
    Width = Main.picMap.ScaleWidth 'get the height and width of the board
    Height = Main.picMap.ScaleHeight
    
    ReDim Board(1 To Width, 1 To Height)
    
    For Y = 1 To Height
        For X = 1 To Width
            a = Main.picMap.Point(X - 1, Y - 1) 'Get the color of this pixel
            Select Case a
            Case 0 'if it's black, do nothing
            Case vbBlue 'If it's blue...
                P1.X = X '... Set player startpoint to here...
                P1.Y = Y
                Board(X, Y) = True '... and mark it as walkable
            Case vbRed 'if it's Red...
                G1.X = X '... move the goal to here...
                G1.Y = Y
                Board(X, Y) = True '... and mark it as walkable
            Case Else
                Board(X, Y) = True 'if it's any other color, mark it as walkable
            End Select
        Next X
    Next Y
End Sub

Public Sub PaintBoard()
    'Paint the Board
    Main.picBuffer.Cls
    For Y = 1 To Height
        For X = 1 To Width
            If Board(X, Y) Then
                BitBlt Main.picBuffer.hDC, (X - 1) * Size, (Y - 1) * Size, Size, Size, Main.picTile.hDC, 0, 0, vbSrcCopy
            End If
        Next X
    Next Y
    
    
    If Not FlagDead Then 'Now paint The Player
        BitBlt Main.picBuffer.hDC, (P1.X - 1) * Size, (P1.Y - 1) * Size, Size, Size, Main.picPM(P1.Dire - 1).hDC, 0, 0, vbSrcAnd
        BitBlt Main.picBuffer.hDC, (P1.X - 1) * Size, (P1.Y - 1) * Size, Size, Size, Main.picP(P1.Dire - 1).hDC, 0, 0, vbSrcPaint
    Else 'The player is dead, paint his scream
        BitBlt Main.picBuffer.hDC, (P1.X - 2) * Size, (P1.Y - 0.7) * Size, Size * 3, Size, Main.picAaaM.hDC, 0, 0, vbSrcAnd
        BitBlt Main.picBuffer.hDC, (P1.X - 2) * Size, (P1.Y - 0.7) * Size, Size * 3, Size, Main.picAaa.hDC, 0, 0, vbSrcPaint
    End If
    'Paint The Flag
    BitBlt Main.picBuffer.hDC, (G1.X - 1) * Size, (G1.Y - 1) * Size, Size, Size, Main.picFlagM(G1.Tag).hDC, 0, 0, vbSrcAnd
    BitBlt Main.picBuffer.hDC, (G1.X - 1) * Size, (G1.Y - 1) * Size, Size, Size, Main.picFlag(G1.Tag).hDC, 0, 0, vbSrcPaint
    
    
    BitBlt Main.picMain.hDC, 0, 0, Main.picMain.ScaleWidth, Main.picMain.ScaleHeight, Main.picBuffer.hDC, 0, 0, vbSrcCopy
End Sub

Public Sub DoKeys()
    
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then
        P1.X = P1.X - 1
        P1.Dire = 1
    End If
    If GetAsyncKeyState(vbKeyDown) <> 0 Then
        P1.Y = P1.Y + 1
        P1.Dire = 2
    End If
    If GetAsyncKeyState(vbKeyRight) <> 0 Then
        P1.X = P1.X + 1
        P1.Dire = 3
    End If
    If GetAsyncKeyState(vbKeyUp) <> 0 Then
        P1.Y = P1.Y - 1
        P1.Dire = 4
    End If
    
    
    If P1.X <= 1 Then P1.X = 1 'Prevent us from going of the edge of the world
    If P1.Y <= 1 Then P1.Y = 1
    If P1.X >= Width Then P1.X = Height
    If P1.Y >= Height Then P1.Y = Width

End Sub

Public Sub CheckFallOff()
    If Board(P1.X, P1.Y) Then Exit Sub 'There is ground under our feet :)
    FlagDead = True 'Falg for our death!
End Sub

Public Sub CheckInGoal()
    If P1.X = G1.X And P1.Y = G1.Y Then 'Are we at our goal?
        EndGame 2 'if so, end the game in style 2: WE WIN! :D
    End If
End Sub
Public Sub EndGame(Why)
    Select Case Why
    Case 1 'Bummer, we lost....
        PlaySound "fall"
        MsgBox "Nope. Wrong! You shouldn't have fallen off!" & vbNewLine _
        & "It's not good for you. Thanks for playing though :)", vbOKOnly, "Game Over"
        End '<-- End ends the game
    Case 2 'Yeay, we won!
        PlaySound "win"
        MsgBox "Hey! You made it! Not too hard I guess..." & vbNewLine _
        & "But thanks for playing anyway! :)", vbOKOnly, "Game Over" '^^^ This _ can be used to break statments into several lines
        End '<-- End ends the game
    End Select
    
End Sub

Public Function PlaySound(File As String)
    wFlags% = SND_ASYNC Or SND_NODEFAULT 'Dont mind this, just copy it if you don't understand :)
    Svar = sndPlaySound(App.Path & "\" & File & ".wav", wFlags%)
End Function
