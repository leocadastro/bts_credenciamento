<%
function IsValidEmail(email)
	isitvalid = true
	dim names, name, i, c
	names = Split(email, "@")
	if UBound(names) <> 1 then
		isitvalid = false
		exit function
	end if
	for each name in names
		if Len(name) <=  0 then
			isitvalid = false
			exit function
		end if
		for i = 1 to Len(name)
			c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
			 isitvalid = false
			 exit function
			end if
		next
		if Left(name, 1) = "." or Right(name, 1) = "." then
			isitvalid = false
			exit function
		end if
	next
	if InStr(names(1), ".") <= 0 then
		isitvalid = false 
		exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
		isitvalid = false
		exit function
	end if
	if InStr(email, "..") > 0 then
		isitvalid = false
	end if
	IsValidEmail = isitvalid
End function
%>