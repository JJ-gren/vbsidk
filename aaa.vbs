Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' --- Configuration ---
strImageUrl = "https://media1.tenor.com/m/HNObdUglgfgAAAAC/skeleton-banging-shield.gif"
strAudioUrl = "https://quicksounds.com/uploads/tracks/515006306_914315829_470386365.mp3"
intWindowCount = 25

' --- Setup Audio ---
Set player = CreateObject("WMPlayer.OCX")
player.URL = strAudioUrl
player.settings.setMode "loop", True
player.settings.volume = 100 
player.Controls.play

' --- Create HTML File ---
strTempFolder = objShell.ExpandEnvironmentStrings("%TEMP%")
strHtmlFile = strTempFolder & "\skeleton_bash.html"

Set objFile = objFSO.CreateTextFile(strHtmlFile, True)
objFile.WriteLine "<html><head><title>Banging Shield</title>"
objFile.WriteLine "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">"
objFile.WriteLine "<style>"
objFile.WriteLine "  body { margin: 0; background-color: black; overflow: hidden; display: flex; align-items: center; justify-content: center; height: 100vh; width: 100vw; }"
objFile.WriteLine "  img { max-width: 100%; max-height: 100%; object-fit: contain; }"
objFile.WriteLine "</style>"
objFile.WriteLine "<script>"
objFile.WriteLine "  document.onkeydown = function(e) {"
objFile.WriteLine "    var ev = e || window.event;"
objFile.WriteLine "  };"
objFile.WriteLine "</script>"
objFile.WriteLine "</head><body><img src=""" & strImageUrl & """></body></html>"
objFile.Close

' --- Launch 3 Windows ---
For i = 1 to intWindowCount
    objShell.Run "mshta.exe """ & strHtmlFile & """ -kiosk", 1, False
    WScript.Sleep 200 ' Short pause to ensure windows initialize
Next

' --- Monitor Processes ---
' Music continues until ALL 3 windows are closed
Do While True
    On Error Resume Next
    Set colProcesses = GetObject("winmgmts:").ExecQuery(_
        "Select * from Win32_Process Where Name = 'mshta.exe' AND CommandLine LIKE '%" & objFSO.GetFileName(strHtmlFile) & "%'")
    
    On Error GoTo 0
    WScript.Sleep 500 
Loop

' Cleanup
On Error Resume Next
objFSO.DeleteFile strHtmlFile
