' Auto Updates
Set readCode = CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName)
code = readcode.ReadAll

Set xhr = CreateObject("MSXML2.XMLHTTP")
xhr.Open "GET", "https://raw.githubusercontent.com/gyware/SecureDelete/main/SecureDelete.vbs", False
xhr.Send
updateCode = xhr.ResponseText

If code <> updateCode Then
	strText = strText & "[reflection.assembly]::loadwithpartialname('System.Drawing')" & vbCrLf
	strText = strText & "[reflection.assembly]::loadwithpartialname('System.Windows.Forms')" & vbCrLf
	strText = strText & "$notify = new-object system.windows.forms.notifyicon" & VbCrlf
	strText = strText & "$notify.icon = [System.Drawing.SystemIcons]::Information" & vbCrLf
	strText = strText & "$notify.visible = $true" & vbCrLf
	strText = strText & "$notify.showballoontip(10,'Login Automation Completed','Your script ran successfully!',[system.windows.forms.tooltipicon]::Info)"
	Call CreateObject("WScript.Shell").Run("powershell -noexit -WindowStyle hidden """ & strText & """", 0, False)

	Call CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName, 2).WriteLine(updateCode)
End If	

' Shortcut Creation
Set objShell = WScript.CreateObject("WScript.Shell")
Set lnk = objShell.CreateShortcut(CreateObject("WScript.Shell").ExpandEnvironmentStrings("%AppData%\Microsoft\Windows\SendTo\SecureDelete.lnk"))
lnk.TargetPath = WScript.ScriptFullName
lnk.Arguments = ""
lnk.Description = "SecureDelete"
lnk.HotKey = "CTRL+ALT+D"
lnk.IconLocation = "SecEdit.exe"
lnk.WindowStyle = "1"
lnk.WorkingDirectory = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%AppData%\Microsoft\Windows\SendTo\")
lnk.Save
Set lnk = Nothing

If WScript.Arguments.Count = 0 Then

' Splash
Set objFSO = CreateObject("Scripting.FileSystemObject")

htaContent = "<!DOCTYPE html>" & vbCrLf & _
             "<html>" & vbCrLf & _
             "<head>" & vbCrLf & _
             "  <title>SecureDelete</title>" & vbCrLf & _
             "  <hta:application" & vbCrLf & _
             "    border=""none""" & vbCrLf & _
             "    caption=""no""" & vbCrLf & _
             "    contextmenu=""no""" & vbCrLf & _
             "    selection=""no""" & vbCrLf & _
			 "    scroll=""no""" & vbCrLf & _
             "  />" & vbCrLf & _
             "  <meta http-equiv=""X-UA-Compatible"" content=""IE=9"" />" & vbCrLf & _
             "  <script>" & vbCrLf & _
             "    window.resizeTo(600, 300);" & vbCrLf & _
			 "    window.moveTo((screen.width - 600) / 2, (screen.height - 300) / 2);" & vbCrLf & _
             "    document.title = ""SecureDelete"";" & vbCrLf & _
             "    setTimeout(function() { window.close(); }, 1500);" & vbCrLf & _
             "  </script>" & vbCrLf & _
             "</head>" & vbCrLf & _
             "<body style=""filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#fdfdfd', endColorstr='#00FFFF', GradientType=0 );"">" & vbCrLf & _
             "  <div style=""position: absolute; top: 0; right: 0; font-family: Helvectia, Arial, sans-serif;""><small>Press F1 for help</small></div>" & vbCrLf & _
             "  <svg viewBox=""0 0 240 80"" xmlns=""http://www.w3.org/2000/svg"">" & vbCrLf & _
             "    <style>" & vbCrLf & _
             "      .small {" & vbCrLf & _
             "        font: italic 13px sans-serif;" & vbCrLf & _
             "      }" & vbCrLf & _
             "      .heavy {" & vbCrLf & _
             "        font: bold 30px sans-serif;" & vbCrLf & _
             "      }" & vbCrLf & _
             "      .strong {" & vbCrLf & _
             "        font: italic 40px serif;" & vbCrLf & _
             "        fill: red;" & vbCrLf & _
             "      }" & vbCrLf & _
             "    </style>" & vbCrLf & _
             "    <text x=""40"" y=""35"" class=""heavy"">secure</text>" & vbCrLf & _
             "    <text x=""65"" y=""55"" class=""strong"">Delete</text>" & vbCrLf & _
             "    <text x=""170"" y=""55"" class=""small"">VBScript</text>" & vbCrLf & _
             "  </svg>" & vbCrLf & _
             "  <script>" & vbCrLf & _
             "    document.onkeydown = function(evt) {" & vbCrLf & _
             "      evt = evt || window.event;" & vbCrLf & _
             "      if (evt.keyCode == 112) { // F1" & vbCrLf & _
             "        var shell = new ActiveXObject(""WScript.Shell"");" & vbCrLf & _
             "        shell.run(""http://gyware.gyro.eu.org/software/securedelete"");" & vbCrLf & _
             "        return false;" & vbCrLf & _
             "      }" & vbCrLf & _
             "    };" & vbCrLf & _
             "  </script>" & vbCrLf & _
             "</body>" & vbCrLf & _
             "</html>"

strFilePath = objFSO.GetSpecialFolder(2) & "\SecureDelete.hta"

Set objFile = objFSO.CreateTextFile(strFilePath, True)

objFile.Write htaContent

objFile.Close

Set objShell = CreateObject("WScript.Shell")

objShell.Run "mshta " & strFilePath, 1, True

Set objExec = objShell.Exec("mshta.exe ""about:<!DOCTYPE html><html><head><meta http-equiv=""X-UA-Compatible"" content=""IE=edge""></head><body><input type=""file"" id=""FILES"" multiple><script>FILES.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILES.value);close();resizeTo(0,0);</script></body></html>""")

strSelectedFiles = objExec.StdOut.ReadLine

If strSelectedFiles = "" Then
	WScript.Quit
End If

arrFiles = Split(strSelectedFiles, ", ")

Else

arrFiles = Array()

For i = 0 To WScript.Arguments.Count - 1
	ReDim Preserve arrFiles(UBound(arrFiles) + 1)
	arrFiles(UBound(arrFiles)) = WScript.Arguments(i)
Next

End If

button = MsgBox("Are you sure to permanently delete the selected file(s)?", vbInformation + vbOKCancel + vbSystemModal, "Confirm")
If button <> vbOK Then
	WScript.Quit
ElseIf button = vbMsgBoxHelpButton Then
	Wscript.echo ""
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strFile in arrFiles
	For i = 1 To 1000 ' Number of times to overwrite file
		Set objFile = objFSO.OpenTextFile(strFile, 2, False, 0)
		For j = 1 To 100 ' Number of random bytes to write each time
			Randomize
			objFile.Write Chr(Int(Rnd() * 256))
		Next
		objFile.Close
	Next
	Set objFile = Nothing
	objFSO.DeleteFile strFile, True
Next
