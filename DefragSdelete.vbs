Set WshShell = WScript.CreateObject("WScript.Shell")
Dim fso, d, dc
Set fso = CreateObject("Scripting.FileSystemObject")
Set dc = fso.Drives
'WshShell.RegWrite "HKCU\Software\Sysinternals\", 0, "REG_SZ"
'WshShell.RegWrite "HKCU\Software\Sysinternals\SDelete\", 0, "REG_SZ"
'WshShell.RegWrite "HKCU\Software\Sysinternals\SDelete\EulaAccepted", 1, "REG_DWORD"
For Each d in dc
If d.DriveType = 2 Then
	Return = WshShell.Run("defrag " & d.DriveLetter & ": /U /X", 1, TRUE)
	Return = WshShell.Run("C:\Repair\sdelete -c -s -z " & d.DriveLetter & ": -accepteula", 1, TRUE)
End If
Next
Set WshShell = Nothing