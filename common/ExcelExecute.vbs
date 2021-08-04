Option Explicit

function OpenExcel(filePath, visibleKbn, readWriteKbn)
	dim rtnValue
	dim objExcelApp
	
	set objExcelApp = WScript.CreateObject("Excel.Application")
	objExcelApp.visible = visibleKbn
	
	if filePath = "" then
		set rtnValue = objExcelApp.workbooks.add
	else
		if readWriteKbn = "r" then
			set rtnValue = objExcelApp.workbooks.open(filePath, false, true)
		else
			set rtnValue = objExcelApp.workbooks.open(filePath, false, false)
		end if
	end if
	
	set OpenExcel = rtnValue
end function

function getOpeningWorkBook(excelName)
	dim objExcelApp
	dim rtnValue
	
	on error resume next
	
	set rtnValue = nothing
	set objExcelApp = GetObject(, "Excel.Application")
	set rtnValue = objExcelApp.workbooks(excelName)
	
	set getOpeningWorkBook = rtnValue
end function


' sample: dim arr(10000, 2)
function TextToExcelWithSplit(wkSheet, startRow, startCol, txtFilePath, charCode, komokuDelim, arr, header)
	dim fileContent, fileContentArray, i, j, lineContentArray, colCnt, headerArray
	
	fileContent = ReadFileAll(txtFilePath, charCode)
	
	colCnt = 0
	if fileContent <> "" then
		fileContentArray = ChangeFileContentToArray(fileContent)
		
		for i = LBound(fileContentArray) to UBound(fileContentArray)
			if fileContentArray(i) <> "" then
				lineContentArray = split(fileContentArray(i), komokuDelim)
				
				colCnt = UBound(lineContentArray) - LBound(lineContentArray)
				
				for j = LBound(lineContentArray) to UBound(lineContentArray)
					arr(i, j) = lineContentArray(j)
				next
			end if
		next
		wkSheet.range(wkSheet.cells(startRow, startCol), newSheet.cells(UBound(fileContentArray) - LBound(fileContentArray) + startRow, colCnt + startCol)) = arr
	end if
	
	if colCnt > 0 and header <> "" and startRow > 1 then
		headerArray = split(header, ",")
		
		if (UBound(headerArray) - LBound(headerArray)) >= colCnt then
			for i = 0 to colCnt
				wkSheet.cells(startRow - 1, startCol + i).value = headerArray(i)
			next
		end if
	end if
	
end function

function TextToExcelWithoutSplit(wkSheet, startRow, startCol, txtFilePath, charCode, objExcelApp)
	dim fileContent, fileContentArray
	
	fileContent = ReadFileAll(txtFilePath, charCode)
	
	if fileContent <> "" then
		fileContentArray = ChangeFileContentToArray(replace(fileContent, chr(9), "    "))
		wkSheet.range(wkSheet.cells(startRow, startCol), newSheet.cells(UBound(fileContentArray) - LBound(fileContentArray) + startRow, startCol)) = objExcelApp.WorksheetFunction.Transpose(fileContentArray)
	end if
end function