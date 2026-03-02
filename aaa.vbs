Dim url, player
url = "https://quicksounds.com/uploads/tracks/515006306_914315829_470386365.mp3"

' Create the Windows Media Player Object
Set player = CreateObject("WMPlayer.OCX")

' Set the URL and enable looping
player.URL = url
player.settings.setMode "loop", True
player.Controls.play

' VBScript is "headless," so we need a loop to keep the script 
' alive while the music plays.
Do While player.playState <> 1 ' 1 = Stopped
    WScript.Sleep 1000 ' Check status every second
Loop