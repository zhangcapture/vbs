Option Explicit

dim commonObjAdodbStream
dim commonObjFileSystem

set commonObjAdodbStream = CreateObject("ADODB.Stream")
set commonObjFileSystem = WScript.CreateObject("Scripting.FileSystemObject")

function OpenFile(filePath, charCode, readWriteKbn, lineSeparator)
	dim objFile

	set objFile = nothing

	if charCode <> "" then
		set objFile = commonObjAdodbStream
		objFile.Type = 2
		objFile.Charset = charCode
		if lineSeparator = "LF" then
			objFile.LineSeparator = 10
		elseif lineSeparator = "CRLF" then
			objFile.LineSeparator = -1
		end if
		objFile.Open
		
		if readWriteKbn = "r" or readWriteKbn = "a" then
			objFile.LoadFromFile filePath
		end if
		
		if readWriteKbn = "a" then
			objFile.Position = objFile.Size
		end if
	else
		if readWriteKbn = "r" then
			Set objFile = commonObjFileSystem.OpenTextFile(filePath, 1)
		elseif readWriteKbn = "w" then
			Set objFile = commonObjFileSystem.OpenTextFile(filePath, 2, true, -2)
		elseif readWriteKbn = "a" then
			Set objFile = commonObjFileSystem.OpenTextFile(filePath, 8, true, -2)
		end if
	end if
	
	set OpenFile = objFile
end function

function ReadFileAll(filePath, charCode)
	dim objFile, fileContent
	
	fileContent = ""
	
	set objFile = OpenFile(filePath, charCode, "r", "")
	
	if charCode <> "" then
		if not objFile.EOS then
			fileContent = objFile.ReadText(-1)
		end if
	else
		if not objFile.AtEndOfStream then
			fileContent = objFile.ReadAll
		end if
	end if
	
	objFile.Close
	set objFile = Nothing
	
	ReadFileAll = fileContent
end function

function IsReadEnd(objFile)
	if TypeName(objFile) = "Stream" then
		IsFileEnd = objFile.EOS
	else
		IsFileEnd = objFile.AtEndOfStream
	end if
end function

function ReadFileLine(objFile)
	if TypeName(objFile) = "Stream" then
		ReadFileLine = objFile.ReadText(-2)
	else
		ReadFileLine = objFile.ReadLine
	end if
end function

function ChangeFileContentToArray(fileContent)
	dim fileContentArray, editedFileContent
	
	editedFileContent = Replace(fileContent, vbcr, "")
	fileContentArray = split(editedFileContent, vblf)
	
	if right(editedFileContent, 1) = vblf then
		ReDim Preserve fileContentArray(Ubound(fileContentArray) - 1)
	end if

	ChangeFileContentToArray = fileContentArray
end function

function FileExists(filePath)
	FileExists = commonObjFileSystem.FileExists(filePath)
end function

function KillFile(filePath)
	commonObjFileSystem.DeleteFile filePath
end function

function ReadFileToArray(filePath, charCode)
	ReadFileToArray = ChangeFileContentToArray(ReadFileAll(filePath, charCode))
end function

sub WriteFile(filePath, charCode, readWriteKbn, fileContent)
	dim objFile

	if charCode <> "" then
		if fileContent <> "" then
			set objFile = OpenFile(filePath, charCode, readWriteKbn, "")
			objFile.WriteText fileContent
			if readWriteKbn = "a" then
				objFile.SaveToFile filePath, 2
			else
				if FileExists(filePath) then
					KillFile(filePath)
				end if
				objFile.SaveToFile filePath, 1
			end if
			objFile.Close
		end if
	else
		set objFile = OpenFile(filePath, charCode, readWriteKbn, "")
		objFile.Write fileContent
		objFile.Close
	end if
	
	set objFile = Nothing
end sub

function WriteEndFileAndMsg(fileName, msg)
	WriteFile fileName, "", "w", ""
	WScript.echo msg
end function

function DeleteNotAvailableCode(fileName)
	dim rtnValue
	dim notAvailableCodes, notAvailableCodeArray, i
	
	rtnValue = fileName
	
	notAvailableCodes = "\/:*?""<>|"
	notAvailableCodeArray = split(notAvailableCodes, "")
	
	for i = LBound(notAvailableCodeArray) to UBound(notAvailableCodeArray)
		rtnValue = replace(rtnValue, notAvailableCodeArray(i), "-")
	next
	
	DeleteNotAvailableCode = rtnValue
end function

function CreateUniqueFilePath(filePath, fileName, fileExt)
	dim rtnValue
	dim editedFilePath, editedFileName, i
	
	editedFilePath = filePath
	if right(editedFilePath, 1) <> "\" then
		editedFilePath = editedFilePath & "\"
	end if
	
	editedFileName = DeleteNotAvailableCode(fileName)
	
	rtnValue = editedFilePath & editedFileName & fileExt
	
	i = 1
	do while FileExists(rtnValue) = true
		rtnValue = editedFilePath & editedFileName & "_" & i & fileExt
		i = i + 1
	loop

	CreateUniqueFilePath = rtnValue
end function

function GetLineDelimiter(fileContent)
	dim rtnValue

	if instr(fileContent, vbcrlf) > 0 then
		rtnValue = vbcrlf
	elseif instr(fileContent, vblf) > 0 then
		rtnValue = vblf
	elseif instr(fileContent, vbcr) > 0 then
		rtnValue = vbcr
	else
		rtnValue = ""
	end if
	
	GetLineDelimiter = rtnValue
end function

function DeleteSpaceLine(fileContent)
	dim rtnValue, strDelims, fileContentArray, i, objRegExp

	strDelims = GetLineDelimiter(fileContent)
	
	if strDelims <> "" then
		set objRegExp = New RegExp
		objRegExp.Global = true
		objRegExp.Pattern = "^\s*$"
	
		fileContentArray = split(fileContent, strDelims)
		
		rtnValue = ""
		for i = LBound(fileContentArray) to UBound(fileContentArray)
			if not objRegExp.Test(fileContentArray(i)) then
				rtnValue = rtnValue & fileContentArray(i) & strDelims
			end if
		next
	else
		rtnValue = fileContent
	end if

	DeleteSpaceLine = rtnValue
end function