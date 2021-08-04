Option Explicit

function GetNowDateTime(rtnType, flgNumberOnly)
	dim nowDateTime, nowDate, nowTime, rtnValue
	
	nowDateTime = Now()
	
	nowDate = Year(nowDateTime) & "/" & LPad(Month(nowDateTime), 2, "0") & "/" & LPad(Day(nowDateTime), 2, "0")
	nowTime = LPad(Hour(nowDateTime), 2, "0") & ":" & LPad(Minute(nowDateTime), 2, "0") & ":" & LPad(Second(nowDateTime), 2, "0")
	
	rtnValue = ""
	if rtnType = "date" then
		rtnValue = nowDate
		
		if flgNumberOnly then
			rtnValue = replace(rtnValue, "/", "")
		end if
	elseif  rtnType = "time" then
		rtnValue = nowTime
		
		if flgNumberOnly then
			rtnValue = replace(rtnValue, ":", "")
		end if
	elseif  rtnType = "datetime" then
		rtnValue = nowDate & " " & nowTime
		
		if flgNumberOnly then
			rtnValue = replace(replace(replace(rtnValue, "/", ""), ":", ""), " ", "")
		end if
	end if
	
	GetNowDateTime = rtnValue
end function

function LPad(src, iLen, padString)
	dim rtnValue, i, editedPadString
	
	rtnValue = src
	if len(src) < iLen then
		if padString = "" then
			editedPadString = " "
		else
			editedPadString = padString
		end if
	
		for i = 1 to (iLen - len(src))
			rtnValue = editedPadString & rtnValue
			
			if len(rtnValue) > iLen then
				rtnValue = right(rtnValue, iLen)
				exit for
			end if
		next
	end if
	
	LPad = rtnValue
end function

function RPad(src, iLen, padString)
	dim rtnValue, i, editedPadString
	
	rtnValue = src
	if len(src) < iLen then
		if padString = "" then
			editedPadString = " "
		else
			editedPadString = padString
		end if
	
		for i = 1 to (iLen - len(src))
			rtnValue = rtnValue & editedPadString
			
			if len(rtnValue) > iLen then
				rtnValue = left(rtnValue, iLen)
				exit for
			end if
		next
	end if
	
	RPad = rtnValue
end function