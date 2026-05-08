Option Explicit
Dim sh, fso, dir, cmdline
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir

cmdline = Chr(34) & dir & "\run_web_background.cmd" & Chr(34)
sh.Run cmdline, 0, False
