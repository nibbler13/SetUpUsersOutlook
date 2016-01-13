set fso = createobject("Scripting.FileSystemObject") 	
startFolder = "\\mssu-fs-01\mail\mail_auto_cfg\"
endFolder = "C:\Temp\"
set objFolder = fso.GetFolder(startFolder)
set colFiles = objFolder.Files
For Each objFile in colFiles
	If Not objFile.Name = "Thumbs.db" Then
		If Not fso.FileExists(endFolder & objFile.Name) Then
			fso.copyfile startFolder & objFile.Name, endFolder
		End If
	End If
Next
Set oWsh = CreateObject("Wscript.Shell")
oWsh.Run("c:\temp\mssu_configurator.exe")