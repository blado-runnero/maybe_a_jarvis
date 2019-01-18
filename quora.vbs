Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
dim str
if hour(time) < 12 then
str = "Good Morning "
else
if hour(time) > 12 then
if hour(time) > 16 then
str = "Good evening "
else
str = "Good afternoon "
end if
end if
end if
str= str & "Welcome Mr. Blade to your own World!"
Sapi.speak str