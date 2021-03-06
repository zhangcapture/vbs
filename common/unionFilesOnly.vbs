Option Explicit

Include CreateObject("WScript.Shell").ExpandEnvironmentStrings("%common%") & "FileReadWrite.vbs", ""

Dim inFilePath01, otFilePath01, inFileCharCode, otFileCharCode
Dim inFileContent

if WScript.Arguments.Count < 4 then
	Wscript.quit
end if

inFilePath01    = WScript.arguments(0)
otFilePath01    = WScript.arguments(1)
inFileCharCode  = WScript.arguments(2)
otFileCharCode  = WScript.arguments(3)

inFileContent = ReadFileAll(inFilePath01, inFileCharCode)

if inFileContent <> "" then
	WriteFile otFilePath01, otFileCharCode, "a", inFileContent
end if

WScript.Quit

function Include(filePath, isUTF8)
	Dim objFSO, objFile

	if isUTF8 = "UTF-8" then
		set objFile = CreateObject("ADODB.Stream")
		objFile.Type = 2
		objFile.Charset = "UTF-8"
		objFile.Open
		objFile.LoadFromFile filePath
		ExecuteGlobal objFile.ReadText(-1)
	else
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(filePath, 1)
		ExecuteGlobal objFile.ReadAll
	end if

	objFile.Close
	Set objFile = Nothing 
	Set objFSO = Nothing
End function
