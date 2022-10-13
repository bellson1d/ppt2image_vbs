' test.vbs

'WScript.Echo "Echo test"  ' doesn't work

'MsgBox "Message box!"     ' look like doesn't work either

' Write to file - works
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("out.txt", True)
objFile.Write "Output to file test" & vbCrLf
objFile.Close