VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flappy Bird"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrChk 
      Interval        =   100
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer tmrRfsh 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim isNight As Boolean, isPlay As Boolean, isStart As Boolean, isOver As Boolean, isHit As Boolean
Dim iFlr As Integer                                   '地面动画变量
'鸟的动画状态，鸟的动画状态变化量，鸟的高度，鸟的初速度v0，鸟的时间t
Dim iBird As Integer, iBirdPlus As Integer, sBirdTop As Single, sBirdV As Single, sBirdT As Single
Const iBirdLeft As Integer = 70
Const sAcc As Single = 0.5                            '鸟的加速度常量
Dim iPipe As Integer, sPipeTop(2) As Single           '管子位置
Const iPipeSpc As Integer = 130                       '管子开口大小及间隔常量
Dim iScr As Integer, iMax As Integer, bPlused(2) As Boolean

Sub PaintImage(sFileName As String, ByVal X As Single, ByVal Y As Single, Optional ByVal cX As Single, Optional ByVal cW As Single, Optional Height As Single)
    If Dir(ResPath & sFileName & ".gif") = "" Then
        MsgBox "资源 """ & ResPath & sFileName & ".gif"" 丢失！", 48, "Flappy Bird VB"
        Unload Me
    End If
    If Dir(ResPath & sFileName & "_m.gif") = "" Then
        If Height = 0 Then
            PaintPicture LoadPicture(ResPath & sFileName & ".gif"), X, Y
        Else
            If Int(Height) <> 0 Then PaintPicture LoadPicture(ResPath & sFileName & ".gif"), X, Y, , Height
        End If
    Else
        If cX = 0 And cW = 0 Then
            PaintPicture LoadPicture(ResPath & sFileName & "_m.gif"), X, Y, , , , , , , vbSrcAnd
            PaintPicture LoadPicture(ResPath & sFileName & ".gif"), X, Y, , , , , , , vbSrcInvert
        Else
            PaintPicture LoadPicture(ResPath & sFileName & "_m.gif"), X, Y, , , cX, , cW, , vbSrcAnd
            PaintPicture LoadPicture(ResPath & sFileName & ".gif"), X, Y, , , cX, , cW, , vbSrcInvert
        End If
    End If
End Sub

Sub PaintScr(ipScr As Integer, Optional isLtl As Boolean = False, Optional isMax As Boolean = False)
    With Me
        Dim sLeft As Single, sWidth As Single, sScr As String, i As Integer
        sScr = CStr(ipScr)
        If isLtl Then
            sWidth = Len(Replace(sScr, "1", "")) * 14 + (Len(sScr) - Len(Replace(sScr, "1", ""))) * 10 + (Len(sScr) - 1) * 2
            sLeft = (.ScaleWidth - 226) / 2 + 226 - sWidth - 22
            Dim iTop As Integer
            iTop = 200 + IIf(isMax, 76, 34)
            For i = 1 To Len(sScr)
                If Mid(sScr, i, 1) = "1" Then
                    PaintImage "Num_1_l", sLeft, iTop
                    sLeft = sLeft + 10 + 2
                Else
                    PaintImage "Nums_l", sLeft, iTop, 14 * Int(Mid(sScr, i, 1)), 14
                    sLeft = sLeft + 14 + 2
                End If
            Next i
        Else
            sWidth = Len(Replace(sScr, "1", "")) * 24 + (Len(sScr) - Len(Replace(sScr, "1", ""))) * 16
            sLeft = (.ScaleWidth - sWidth) / 2
            For i = 1 To Len(sScr)
                If Mid(sScr, i, 1) = "1" Then
                    PaintImage "Num_1_b", sLeft, 50
                    sLeft = sLeft + 16
                Else
                    PaintImage "Nums_b", sLeft, 50, 24 * Int(Mid(sScr, i, 1)), 24
                    sLeft = sLeft + 24
                End If
            Next i
        End If
    End With
End Sub

Function GetPipeTop() As Single
    GetPipeTop = 26 + (Me.ScaleHeight - 26 * 2 - 112 - iPipeSpc) * Rnd
End Function

Sub InitGame()
    With Me
        .Cls
        .Picture = LoadPicture("")
        isNight = Int(Rnd * 3) < 1
        DrawBg
        PaintImage "Logo", (.ScaleWidth - 178) / 2, 130
        PaintImage "Play", (.ScaleWidth - 104) / 2, 330
        .Picture = .Image
        isPlay = False
        isStart = False
        isOver = False
        isHit = False
        bPlused(0) = False
        bPlused(1) = False
        bPlused(2) = False
        iScr = 0                                '初始化分数
        sBirdV = -6                             '初始化鸟的物理量
        sBirdTop = 200
        iPipe = 500                             '初始化管子位置
        sPipeTop(0) = GetPipeTop
        sPipeTop(1) = GetPipeTop
        sPipeTop(2) = GetPipeTop
   End With
End Sub

Sub DrawBg()
    If isNight Then
        PaintImage "Bg_Night", 0, 0
        PaintImage "Bg_Night", 288, 0
    Else
        PaintImage "Bg_Day", 0, 0
        PaintImage "Bg_Day", 288, 0
    End If
End Sub

Sub DrawReady()
    With Me
        .Cls
        .Picture = LoadPicture("")
        DrawBg
        PaintImage "Ready", (.ScaleWidth - 184) / 2, 130
        PaintImage "Tap", (.ScaleWidth - 114) / 2, 200
        .Picture = .Image
    End With
End Sub

Sub DrawPipes()
    Dim i As Integer
    For i = 0 To 2
        PaintImage "Pipe_U", iPipe + (iPipeSpc + 30) * i, sPipeTop(i)
        PaintImage "Pipe_L", iPipe + (iPipeSpc + 30) * i + 2, 0, , , sPipeTop(i)
        PaintImage "Pipe_D", iPipe + (iPipeSpc + 30) * i, sPipeTop(i) + iPipeSpc
        PaintImage "Pipe_L", iPipe + (iPipeSpc + 30) * i + 2, sPipeTop(i) + iPipeSpc + 26, , , Me.ScaleHeight - 112 - (sPipeTop(i) + iPipeSpc + 26)
        If iPipe + (iPipeSpc + 30) * i <= iBirdLeft + 33 And Not bPlused(i) Then
            iScr = iScr + 1
            'PlaySfx SFX_POINT
            frmIcon.PlayPoint
            bPlused(i) = True
        End If
        If iPipe + (iPipeSpc + 30) * i <= iBirdLeft + 34 And iPipe + (iPipeSpc + 30) * i + 52 >= iBirdLeft And Not isHit Then
            If sBirdTop <= sPipeTop(i) + 26 Or sBirdTop >= sPipeTop(i) + iPipeSpc - 24 Then
                PlaySfx SFX_DIE
                isHit = True
            End If
        End If
    Next i
End Sub

Sub DrawOver()
    With Me
        Dim bNew As Boolean
        bNew = False
        If Not isHit Then PlaySfx SFX_HIT
        If iScr > iMax Then
            iMax = iScr                                            '设置最高分数
            bNew = True
        End If
        PaintImage "GameOver", (.ScaleWidth - 192) / 2, 130
        PaintImage "Board", (.ScaleWidth - 226) / 2, 200
        Select Case Int(iScr / 10) - 1
            Case 0 To 3
                PaintImage "Medals", (.ScaleWidth - 226) / 2 + 26, 200 + 42, 44 * (Int(iScr / 10) - 1), 44
            Case Else
                If Int(iScr / 10) - 1 <> -1 Then PaintImage "Medals", (.ScaleWidth - 226) / 2 + 26, 200 + 42, 44 * 3, 44
        End Select
        PaintScr iScr, True
        PaintScr iMax, True, True
        If bNew Then PaintImage "New", (.ScaleWidth - 226) / 2 + 226 - 32 - 60, 258
        PaintImage "Play", (.ScaleWidth - 104) / 2, 330
        .Picture = .Image
    End With
End Sub

Private Sub Form_Load()
    ResPath = App.Path & "\Res\"
    SfxPath = App.Path & "\Sfx\"
    If App.PrevInstance Then End
    Load frmIcon
    iScr = 0
    iMax = Int(GetSetting("FlapVB", "Data", "Score", "0"))
    Randomize
    InitGame
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not isHit Then
        If isOver Then
            If X > (Me.ScaleWidth - 104) / 2 And X < (Me.ScaleWidth - 104) / 2 + 104 And Y > 330 And Y < 388 Then    '点击了 Play
                InitGame
                PlaySfx SFX_SWOOSH
            End If
        ElseIf isPlay Then
            If Not isStart Then
                Me.Cls
                DrawBg
                Me.Picture = Me.Image
                isStart = True
            End If
            sBirdT = 0
            PlaySfx SFX_WING
        Else
            If X > (Me.ScaleWidth - 104) / 2 And X < (Me.ScaleWidth - 104) / 2 + 104 And Y > 330 And Y < 388 Then    '点击了 Play
                isPlay = True
                DrawReady
                PlaySfx SFX_SWOOSH
            End If
        End If
    ElseIf Button = 1 And isHit And isOver Then
        If X > (Me.ScaleWidth - 104) / 2 And X < (Me.ScaleWidth - 104) / 2 + 104 And Y > 330 And Y < 388 Then        '点击了 Play
            InitGame
            PlaySfx SFX_SWOOSH
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmIcon
    SaveSetting "FlapVB", "Data", "Score", iMax
End Sub

Private Sub tmrChk_Timer()                                                    '判断焦点
    If GetForegroundWindow <> Me.hWnd Then
        If tmrRfsh.Enabled Then tmrRfsh.Enabled = False
    Else
        If Not tmrRfsh.Enabled Then tmrRfsh.Enabled = True
    End If
End Sub

Private Sub tmrRfsh_Timer()
    With Me
        .Cls
        
        '画地面
        If (Not isHit) And (Not isOver) Then
            iFlr = iFlr - 2
            If iFlr < -336 Then iFlr = 0
        End If
        PaintImage "Floor", iFlr, .ScaleHeight - 112
        PaintImage "Floor", iFlr + 336, .ScaleHeight - 112
        PaintImage "Floor", iFlr + 336 * 2, .ScaleHeight - 112
        
        '画版权信息
        If Not isPlay Then PaintImage "Copyright", (.ScaleWidth - 122) / 2, .ScaleHeight - 80
        
        '画管子
        If isStart And (Not isHit) And (Not isOver) Then
            iPipe = iPipe - 2
            If iPipe <= -52 Then
                iPipe = iPipeSpc + 30 - 52
                sPipeTop(0) = sPipeTop(1)
                sPipeTop(1) = sPipeTop(2)
                sPipeTop(2) = GetPipeTop
                bPlused(0) = bPlused(1)
                bPlused(1) = bPlused(2)
                bPlused(2) = False
            End If
        End If
        If Not isOver Then DrawPipes
        
        '画个鸟
        If Not isOver Then
            If iBird = 10 Then iBirdPlus = -1                      '鸟的动画
            If iBird = 0 Then iBirdPlus = 1
            iBird = iBird + iBirdPlus
        End If
        If isOver Then
            PaintImage "Bird", iBirdLeft, .ScaleHeight - 112 - 24, 34, 34
        ElseIf isStart Then
            sBirdT = sBirdT + 0.7
            sBirdTop = sBirdTop + sBirdV + sBirdT * sAcc
            If sBirdTop < -30 Then sBirdTop = -30                  '别飞太高
            If sBirdTop >= .ScaleHeight - 112 - 24 Then            '撞到地面
                isOver = True
                PaintImage "Bird", iBirdLeft, .ScaleHeight - 112 - 24, 34, 34
                DrawOver
            Else
                PaintImage "Bird", iBirdLeft, sBirdTop, 34 * Int(iBird / 4.5), 34
            End If
        ElseIf isPlay Then
            PaintImage "Bird", iBirdLeft, 200 + iBird, 34 * Int(iBird / 4.5), 34
        Else
            PaintImage "Bird", (.ScaleWidth - 34) / 2, 200 + iBird, 34 * Int(iBird / 4.5), 34
        End If
        
        '画分数
        If isPlay And Not isOver Then PaintScr iScr
        
    End With
End Sub
