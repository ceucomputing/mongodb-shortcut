' Adapted from https://gist.github.com/rcmdnk/7631748

' Set basic objects
Set wssh = WScript.CreateObject("WScript.Shell")
Set fs = WScript.CreateObject("Scripting.FileSystemObject")

' Basic Values
progFiles = "C:\Program Files\MongoDB\Server\3.4\bin"
startMenu = wssh.SpecialFolders("AllUsersPrograms")
pExe = "mongod.exe"

' Make start menu shortcut
Set oMyShortCut = wssh.CreateShortcut(startMenu & "\MongoDB Server.lnk")
oMyShortCut.TargetPath = progFiles & "\" & pExe
oMyShortCut.Arguments = "--dbpath ""%TEMP%"""
oMyShortCut.Save
