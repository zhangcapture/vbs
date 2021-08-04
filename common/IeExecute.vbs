Option Explicit 

function getIEDoc(strTitle, matchType, isWait, rtnType)
	dim  rtnValue, objShell, objWindow

	set rtnValue = nothing
	
	set objShell = WScript.CreateObject("Shell.Application")
	for each objWindow in objShell.Windows
		if objWindow.name = "Internet Explorer" then
			if matchType = "full" and objWindow.document.title = strTitle then
				set rtnValue = objWindow
				exit for
			elseif matchType = "part" and instr(objWindow.document.title, strTitle) > 0 then
				set rtnValue = objWindow
				exit for
			end if
		end if
	next
	
	if not rtnValue is nothing then
		if isWait = "wait" then
			do while rtnValue.busy = true or rtnValue.readyState <> 4
				wscript.sleep 1000
			loop
		end if
		
		if rtnType <> "IE" then
			set rtnValue = rtnValue.document
		end if
	end if
	
	set objShell = nothing
	
	set getIEDoc = rtnValue
end function

function getIEByUrl(strUrl, matchType, isWait, rtnType)
	dim  rtnValue, objShell, objWindow

	set rtnValue = nothing
	
	set objShell = WScript.CreateObject("Shell.Application")
	for each objWindow in objShell.Windows
		if objWindow.name = "Internet Explorer" then
			if matchType = "full" and objWindow.LocationURL = strUrl then
				set rtnValue = objWindow
				exit for
			elseif matchType = "part" and instr(objWindow.LocationURL, strUrl) > 0 then
				set rtnValue = objWindow
				exit for
			end if
		end if
	next
	
	if not rtnValue is nothing then
		if isWait = "wait" then
			do while rtnValue.busy = true or rtnValue.readyState <> 4
				wscript.sleep 1000
			loop
		end if
		
		if rtnType <> "IE" then
			set rtnValue = rtnValue.document
		end if
	end if
	
	set objShell = nothing
	
	set getIEByUrl = rtnValue
end function

function createIEApp()
	dim rtnValue, tryCount, i
	
	on error resume next
	
	tryCount = 3
	set rtnValue = nothing
	
	for i = 1 to tryCount
		set rtnValue = WScript.CreateObject("InternetExplorer.Application")
		
		if not rtnValue is nothing then
			exit for
		end if
	next
	
	set createIEApp = rtnValue
end function

function getFirstLineFromFile(filePath, isUTF8, lineSeparator)
	Dim objFSO, objFile, rtnValue

	rtnValue = ""
	if isUTF8 = "UTF-8" then
		set objFile = CreateObject("ADODB.Stream")
		objFile.Type = 2
		objFile.Charset = "UTF-8"
		if lineSeparator = "LF" then
			objFile.LineSeparator = 10
		elseif lineSeparator = "CRLF" then
			objFile.LineSeparator = -1
		end if
		objFile.Open
		objFile.LoadFromFile filePath
		
		if not objFile.EOS then
			rtnValue = objFile.ReadText(-2)
		end if
	else
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.OpenTextFile(filePath, 1)
		if not objFile.AtEndOfStream then
			rtnValue = objFile.ReadLine
		end if
	end if

	objFile.Close

	Set objFile = Nothing 
	Set objFSO = Nothing
	
	getFirstLineFromFile = rtnValue
end function

function jqueryText(dom, selector, flgDelCrLf)
	dim rtnValue
	dim domSelected
	
	set domSelected = jquery(dom, selector)
	
	rtnValue = ""
	if not domSelected is nothing then
		rtnValue = domSelected.innerText
		if flgDelCrLf then
			rtnValue = delCrLf(rtnValue)
		end if
	end if

	jqueryText = rtnValue
end function

function jquery(dom, selector)
	dim rtnValue
	dim selectorKbn, selectorContent
	dim arraySelector, i
	
	set rtnValue = nothing
	
	if instr(selector, " ") > 0 then
		arraySelector = split(selector, " ")
		
		set rtnValue = dom
		for i = Lbound(arraySelector) to Ubound(arraySelector)
			if arraySelector(i) <> "" then
				set rtnValue = jquery(rtnValue, arraySelector(i))
			end if
		next
	else
		selectorKbn = left(selector, 1)
		selectorContent = mid(selector, 2)
		
		if selectorKbn = "#" then
			set rtnValue = jqueryId(dom, selectorContent)
		elseif selectorKbn = "." then
			set rtnValue = jqueryClassName(dom, selectorContent)
		else
			set rtnValue = jqueryTagName(dom, selector)
		end if
	end if

	set jquery = rtnValue
end function

function jqueryId(dom, id)
	dim rtnValue
	
	set rtnValue = nothing
	
	if not dom is nothing then
		on error resume next
		set rtnValue = dom.getElementById(id)
	end if
	
	set jqueryId = rtnValue
end function

function jqueryClassName(dom, className)
	dim rtnValue
	dim startOfSeparator
	dim classNameOnly, indexNo
	
	set rtnValue = nothing
	
	if not dom is nothing then
		startOfSeparator = instr(className, ":")
		if startOfSeparator > 0 then
			classNameOnly = mid(className, 1, startOfSeparator - 1)
			indexNo = CInt(mid(className, startOfSeparator + 1))
		else
			classNameOnly = className
			indexNo = -1
		end if
		
		set rtnValue = dom.getElementsByClassName(classNameOnly)
		
		if rtnValue.length = 0 then
			set rtnValue = nothing
		else
			if indexNo > -1 then
				if rtnValue.length > indexNo then
					set rtnValue = rtnValue(indexNo)
				else
					set rtnValue = nothing
				end if
			end if
		end if
	end if
	
	set jqueryClassName = rtnValue
end function

function jqueryTagName(dom, tagName)
	dim rtnValue
	dim startOfSeparator
	dim tagNameOnly, indexNo
	
	set rtnValue = nothing
	
	if not dom is nothing then
		startOfSeparator = instr(tagName, ":")
		if startOfSeparator > 0 then
			tagNameOnly = mid(tagName, 1, startOfSeparator - 1)
			indexNo = CInt(mid(tagName, startOfSeparator + 1))
		else
			tagNameOnly = tagName
			indexNo = -1
		end if
	
		set rtnValue = dom.getElementsByTagName(tagNameOnly)
		
		if rtnValue.length = 0 then
			set rtnValue = nothing
		else
			if indexNo > -1 then
				if rtnValue.length > indexNo then
					set rtnValue = rtnValue(indexNo)
				else
					set rtnValue = nothing
				end if
			end if
		end if
	end if
	
	set jqueryTagName = rtnValue
end function

function jqueryTagClass(dom, tagClassName)
	dim rtnValue
	dim tagName, className, tagClassNameOnly, indexNo, startOfSeparator
	dim domClass, i, cnt, matchFlg
	
	set rtnValue = nothing
	
	if not dom is nothing then
		startOfSeparator = instr(tagClassName, ":")
		if startOfSeparator > 0 then
			tagClassNameOnly = mid(tagClassName, 1, startOfSeparator - 1)
			indexNo = CInt(mid(tagClassName, startOfSeparator + 1))
		else
			tagClassNameOnly = tagClassName
			indexNo = -1
		end if
	
		startOfSeparator = instr(tagClassNameOnly, ".")
		tagName = UCase(mid(tagClassNameOnly, 1, startOfSeparator - 1))
		className = mid(tagClassNameOnly, startOfSeparator + 1)

		set rtnValue = jqueryClassName(dom, className)
		
		if not rtnValue is nothing then
			cnt = 0
			matchFlg = false
			for i = rtnValue.length - 1 to 0 step -1
				if UCase(rtnValue(i).tagName) = tagName then
					if indexNo = cnt then
						matchFlg = true
						set rtnValue = rtnValue(i)
						exit for
					end if
					cnt = cnt + 1
				end if
			next
			
			if not matchFlg then
				set rtnValue = nothing
			end if
		end if
	end if
	
	set jqueryTagClass = rtnValue
end function

function delCrLf(fromStr)
	dim rtnValue
	
	rtnValue = fromStr
	rtnValue = replace(rtnValue, vbcr, "")
	rtnValue = replace(rtnValue, vblf, "")

	delCrLf = rtnValue
end function
