Set WshShell = WScript.CreateObject("WScript.Shell")
Dim fso, d, dc
Set fso = CreateObject("Scripting.FileSystemObject")
Set dc = fso.Drives
For Each d in dc
If d.DriveType = 2 Then
	Return = WshShell.Run("defrag " & d.DriveLetter & ": /U /X", 1, TRUE)
	Return = WshShell.Run("C:\Repair\sdelete -c -s -z " & d.DriveLetter & ": -accepteula", 1, TRUE)
End If
Next
Set WshShell = Nothing