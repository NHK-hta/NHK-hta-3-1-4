option explicit

function get_base_date(n)
	dim retval,y,m,d
	dim w
	y = Year(n)
	m = Month(n)
	if m <= 3 then
		y = y - 1
	end if
	retval = DateValue(CStr(y) & "/4/1")
	w = Weekday(retval)
	if w = vbSaturday then	'土曜日だったら日曜日
		retval = DateAdd("d",1,retval)
	elseif w = vbSunday then	'日曜だったら何もしない
	else
		retval = DateAdd("d",-w+1,retval)
	end if
	d=get_diff_day(n,retval)
	if (d < 8) or ( (d=8) and (Hour(n)<10) ) then
		retval = DateValue(CStr(y-1) & "/4/1")
		w = Weekday(retval)
		if w = vbSaturday then	'土曜日だったら日曜日
			retval = DateAdd("d",1,retval)
		elseif w = vbSunday then	'日曜だったら何もしない
		else
			retval = DateAdd("d",-w+1,retval)
		end if
	end if
	get_base_date = retval
end function

function get_diff_day(d,base_d)
	dim df,retval
	df = get_date_str(d)
	retval = DateDiff("d",base_d,df)
	do while (retval > 356) and (Month(df)=12)
		df = DateAdd("yyyy",-1,df)
		retval = DateDiff("d",base_d,df)
	loop
	get_diff_day = retval
end function

function get_last_year(d,base_d)
	dim df,diff
	df = get_date_str(d)
	diff = DateDiff("d",base_d,df)
	do while (diff > 356) and (Month(df)=12)
		df = DateAdd("yyyy",-1,df)
		diff = DateDiff("d",base_d,df)
	loop
	get_last_year = Year(df)
end function

function get_date_str(d)
	Dim regEx
	Set regEx = New RegExp   ' 正規表現を作成します。
	dim Match, Matches,da,df
	regEx.IgnoreCase = True
	regEx.Global = False
	regEx.Pattern = "[0-9][0-9]*年[0-9][0-9]*月[0-9][0-9]*日"
	Set Matches = regEx.Execute(d)
	da = ""
	For Each Match in Matches
		da=Match
	Next
	if da = "" then
		regEx.Pattern = "[0-9][0-9]*月[0-9][0-9]*日"
		Set Matches = regEx.Execute(d)
		For Each Match in Matches
			da=Match
		Next
		if da = "" then
			regEx.Pattern = "[0-9][0-9]*/[0-9][0-9]*/[0-9][0-9]*"
			Set Matches = regEx.Execute(d)
			For Each Match in Matches
				da=Match
			Next
			if da = "" then
				regEx.Pattern = "[0-9][0-9]*/[0-9][0-9]*"
				Set Matches = regEx.Execute(d)
				For Each Match in Matches
					da=Match
				Next
				if da = "" then
					msgbox "警告 : 放送日のフォーマットが違います : " & d
					da = date()
				end if
			end if
		end if
	end if
	df = DateValue(da)
	get_date_str = df
	Set regEx = nothing
end function
