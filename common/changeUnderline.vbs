Option Explicit

dim inputVar
dim rtnValue
dim inputVarArray
dim i

if WScript.Arguments.Count < 1 then
	Wscript.quit
end if

inputVar    = WScript.arguments(0)

rtnValue = ""
inputVarArray = split(inputVar, "_")

for i = LBound(inputVarArray) to UBound(inputVarArray)
	if inputVarArray(i) <> "" then
		if len(inputVarArray(i)) = 1 then
			rtnValue = rtnValue & UCase(inputVarArray(i))
		else
			rtnValue = rtnValue & UCase(left(inputVarArray(i), 1)) & mid(inputVarArray(i), 2)
		end if
	end if
next

wscript.echo rtnValue