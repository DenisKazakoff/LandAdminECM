Option Explicit
 
Dim wshShell
  
Set wshShell = CreateObject( "WScript.Shell" )
wshShell.Run Trim( "launcher.cmd" ), 0, False
Set wshShell = Nothing
