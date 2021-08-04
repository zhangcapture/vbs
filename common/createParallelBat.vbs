Option Explicit
Include CreateObject("WScript.Shell").ExpandEnvironmentStrings("%common%") & "FileReadWrite.vbs", ""

dim template, jobName, commitInterval, maxCommitedCnt, partsBefore, partsMain, partsAfter
dim templateContent, outContent, hashParts, flgChanged, partsKey

if WScript.Arguments.Count < 7 then
	Wscript.quit
end if

template            = WScript.arguments(0)
jobName             = WScript.arguments(1)
commitInterval      = WScript.arguments(2)
maxCommitedCnt      = WScript.arguments(3)
partsBefore         = WScript.arguments(4)
partsMain           = WScript.arguments(5)
partsAfter          = WScript.arguments(6)

templateContent = ReadFileAll(template, "UTF-8")
outContent = templateContent

outContent = replace(outContent, "{commitInterval}", commitInterval)
outContent = replace(outContent, "{maxCommitedCnt}", maxCommitedCnt)

set hashParts = CreateObject("Scripting.Dictionary")
hashParts.Add "{before}", partsBefore
hashParts.Add "{main}",   partsMain
hashParts.Add "{after}",  partsAfter

for each partsKey in hashParts
	flgChanged = false
	if hashParts(partsKey) <> "" then
		if FileExists(hashParts(partsKey)) then
			outContent = replace(outContent, partsKey, ReadFileAll(hashParts(partsKey), "UTF-8"))
			flgChanged = true
		end if
	end if
	if not flgChanged then
		outContent = replace(outContent, partsKey, "")
	end if
next

set hashParts = nothing

WriteFile jobName, "SJIS", "w", outContent
Wscript.quit

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