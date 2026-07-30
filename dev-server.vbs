' Silently start the Pi dev server without showing any window
' Get the directory where this .vbs file lives
Set objShell = CreateObject("Wscript.Shell")
scriptDir = objShell.CurrentDirectory
batPath = scriptDir & "\dev-server.bat"
objShell.Run "cmd /c """ & batPath & """", 0, False
