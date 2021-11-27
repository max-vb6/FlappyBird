Attribute VB_Name = "modMian"
Option Explicit

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC As Long = &H1   ' play asynchronously
Public Const SND_FILENAME As Long = &H20000   ' name is a file name

Public ResPath As String, SfxPath As String

Public Const SFX_DIE As Long = 0
Public Const SFX_HIT As Long = 1
Public Const SFX_POINT As Long = 2
Public Const SFX_SWOOSH As Long = 3
Public Const SFX_WING As Long = 4

Public Sub PlaySfx(lSfx As Long)
    Dim sSfx() As String
    sSfx = Split("die,hit,point,swooshing,wing", ",")
    If Dir(SfxPath & "sfx_" & sSfx(lSfx) & ".wav") = "" Then
        MsgBox "ÉùÒô """ & SfxPath & "sfx_" & sSfx(lSfx) & ".wav"" ¶ªÊ§£¡", 48, "Flappy Bird VB"
        Unload frmMain
        Exit Sub
    End If
    PlaySound SfxPath & "sfx_" & sSfx(lSfx) & ".wav", 0, SND_FILENAME Or SND_ASYNC
End Sub
