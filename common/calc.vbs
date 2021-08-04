Option Explicit

dim operator, num1, num2, decimalLen
dim rtnValue

if WScript.Arguments.Count < 4 then
	Wscript.quit
end if

operator    = WScript.arguments(0)
num1        = WScript.arguments(1)
num2        = WScript.arguments(2)
decimalLen  = WScript.arguments(3)

if operator = "minus" then
	rtnValue = Round(num1 - num2, decimalLen)
elseif operator = "plus" then
	rtnValue = Round(num1 + num2, decimalLen)
end if

if left(CStr(rtnValue), 1) = "." then
	rtnValue = "0" & rtnValue
end if

wscript.echo rtnValue