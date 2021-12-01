option explicit

dim	ini_count,ini_change
dim	ini_buff(),ini_section(),ini_key(),ini_value()
redim	ini_buff(50),ini_section(50),ini_key(50),ini_value(50)

sub flag_read(arg_section,arg_key,arg_name,flag,flag_default)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	dim	str,num
	if ini_search(arg_section,arg_key,str,num) then
		regEx.IgnoreCase = True
		regEx.Global = False
		regEx.Pattern = "true"
		if regEx.Test(str) then
			flag = CBool(str)
		else
			regEx.Pattern = "false"
			if regEx.Test(str) then
				flag = CBool(str)
			else
				msgbox "ini-file の " & arg_name & " 設定が不明:" & str
				flag = flag_default
				ini_add arg_section, arg_key, CStr(flag)
			end if
		end if
	else
		'msgbox "ini-file に " & arg_name & " 設定が無い"
		flag = flag_default
		ini_add arg_section, arg_key, CStr(flag)
	end if

	set regEx = nothing
end sub

sub str_read(arg_section, arg_key, arg_name, arg_str, str_default)
	dim	str,num
	dim	f(100),nf
	if ini_search(arg_section,arg_key,str,num) then
		split_comma str,f,nf
		if (nf >= 1) and (f(1)<>"") then
			arg_str = f(1)
		else
			'msgbox "ini-file の " & arg_name & " 設定が不明:" & str
			arg_str = str_default
			ini_add arg_section, arg_key, """" & arg_str & """"
		end if
	else
		'msgbox "ini-file に " & arg_name & " 設定が無い"
		arg_str = str_default
		ini_add arg_section, arg_key, """" & arg_str & """"
	end if
end sub


function read_ini_file
	dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	ini_count=0
	ini_buff(0)=""
	ini_section(0)=""
	ini_key(0)=""
	ini_value(0)=""

	if not fso.FileExists(ini_file) then
		read_ini_file=0
		exit function
	end if

	dim ini_in
	Const ForReading=1
	On Error Resume Next
	Set ini_in = fso.OpenTextFile(ini_file, ForReading)
	if Err.Number <> 0 then
		MsgBox "ファイル """ & ini_file & """ が読み込み用に開けません" & vbNewLine & _
				"エラー番号 " & CStr(Err.Number) & vbNewLine & Err.Description
		On Error Goto 0
		Window.Close
	end if
	On Error Goto 0

	ini_change=false

	dim cur_section
	cur_section=""
	Do Until ini_in.AtEndOfStream
		ini_count = ini_count + 1
		if UBound(ini_buff) <= ini_count then
			redim_ini
		end if
		ini_buff(ini_count) = ini_in.ReadLine
		regEx.Pattern = "^\[.*\]$"
		regEx.IgnoreCase = True
		regEx.Global = False
		if regEx.Test(ini_buff(ini_count)) then
			cur_section=ini_buff(ini_count)
			ini_section(ini_count) = cur_section
			ini_key(ini_count) = ini_buff(ini_count)
			ini_value(ini_count)=""
		else
			regEx.IgnoreCase = True
			regEx.Global = False
			regEx.Pattern = "^[ 	]*;"
			if regEx.Test(ini_buff(ini_count)) then
				ini_section(ini_count) = cur_section
				ini_key(ini_count) = ";"
				ini_value(ini_count)= ini_buff(ini_count)
			else
				regEx.IgnoreCase = False
				regEx.Global = False
				regEx.Pattern = "^[ 	]*[a-zA-Z][a-zA-Z0-9_]*[ 	]*="
				if regEx.Test(ini_buff(ini_count)) then
					ini_section(ini_count) = cur_section

					regEx.Pattern = "=.*"
					ini_key(ini_count) = regEx.Replace(ini_buff(ini_count),"")
					regEx.Pattern = "^[ 	]*"
					ini_key(ini_count) = regEx.Replace(ini_key(ini_count),"")
					regEx.Pattern = "[ 	]*$"
					ini_key(ini_count) = regEx.Replace(ini_key(ini_count),"")

					regEx.Pattern = "[^=]*="
					ini_value(ini_count) = regEx.Replace(ini_buff(ini_count),"")
					regEx.Pattern = "^[ 	]*"
					ini_value(ini_count) = regEx.Replace(ini_value(ini_count),"")
					regEx.Pattern = "[ 	]*$"
					ini_value(ini_count) = regEx.Replace(ini_value(ini_count),"")
				else
					msgbox ini_file & " 異常なデータです" & vbNewLine & ini_buff(ini_count)
					ini_count = ini_count - 1
					ini_change = true
				end if
			end if
		end if
	Loop
	ini_in.Close

	read_ini_file = ini_count

	set regEx = nothing
	set fso = nothing
end function


sub write_ini_file
	dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	if not ini_change then
		exit sub
	end if
	dim ini_out
	Const ForWriting = 2
	On Error Resume Next
	Set ini_out = fso.OpenTextFile(ini_file, ForWriting, True)
	if Err.Number <> 0 then
		MsgBox "ファイル """ & ini_file & """ が書き込み用に開けません" & vbNewLine & _
				"エラー番号 " & CStr(Err.Number) & vbNewLine & Err.Description
		On Error Goto 0
		exit sub
	end if
	On Error Goto 0
	dim i
	for i = 1 to ini_count
		ini_out.WriteLine ini_buff(i)
	next
	ini_out.Close
	ini_change = false
	set fso = nothing
end sub

sub redim_ini
	dim newlen
	newlen=UBound(ini_buff)+40
	redim Preserve ini_buff(newlen),ini_section(newlen)
	redim Preserve ini_key(newlen),ini_value(newlen)
end sub

function ini_search(arg_section, arg_key, arg_value, arg_num)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	if ini_count = 0 then
		arg_value = ""
		arg_num = 0
		ini_search = false
		exit function
	end if
	regEx.IgnoreCase = True
	regEx.Global = False
	regEx.Pattern = "^\[.*\]$"
	if (arg_section<>"") and (not regEx.Test(arg_section)) then
		msgbox "In ini_search : section の入力が異常 :" & arg_section
		arg_value = ""
		arg_num = 0
		ini_search = false
		exit function
	end if
	if (arg_key="") then
		msgbox "In ini_search : key の入力が異常 :" & arg_key
		arg_value = ""
		arg_num = 0
		ini_search = false
		exit function
	end if
	dim i, section_hit
	section_hit = false
	for i = 0 to ini_count
		if arg_section = ini_section(i) then
			section_hit = true
			if arg_key = ini_key(i) then
				ini_search = true
				arg_value = ini_value(i)
				arg_num = i
				exit function
			end if
		elseif section_hit then
			ini_search = false
			arg_value = ""
			arg_num = i-1
			exit function
		end if
	next
	ini_search = false
	arg_value = ""
	arg_num = ini_count

	set regEx = nothing
end function

sub ini_add(arg_section, arg_key, arg_value)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	regEx.IgnoreCase = false
	regEx.Global = False
	regEx.Pattern = "^\[.*\]$"
	if (arg_section<>"") and (not regEx.Test(arg_section)) then
		msgbox "In ini_add : section の入力が異常 :" & arg_section
		exit sub
	end if

	dim str, num

	if ini_search(arg_section, arg_key, str, num) then
		if arg_key = ";" then
			ini_add_comment num,arg_value
			ini_change = true
		elseif str <> arg_value then
			ini_value(num) = arg_value
			ini_buff(num) = arg_key & "=" & arg_value
			ini_change = true
		end if
	else
		num = num + 1
		if arg_section <> ini_section(num-1) then
			ini_add_section num,arg_section
			num = num + 1
			ini_change = true
		end if
		if arg_key = ";" then
			ini_add_comment num,arg_value
			ini_change = true
		else
			ini_insert(num)
			ini_section(num) = arg_section
			ini_key(num) = arg_key
			ini_value(num) = arg_value
			ini_buff(num) = arg_key & "=" & arg_value
			ini_change = true
		end if
	end if
	set regEx = nothing
end sub

sub ini_add_comment(num,arg_value)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	if num > ini_count+1 then
		msgbox "ini_add_comment : num が大きすぎる"
		num = ini_count + 1
	elseif num<0 then
		num = 0
	end if
	do while (ini_count>=num) and (ini_key(num)=";")
		num = num + 1
	loop
	ini_insert(num)
	regEx.IgnoreCase = false
	regEx.Global = False
	regEx.Pattern = "^[ 	]*;"
	if regEx.Test(arg_value) then
		ini_buff(num) = arg_value
	else
		ini_buff(num) = ";	" & arg_value
	end if
	ini_change = true
	set regEx = nothing
end sub

sub ini_add_section(num,arg_section)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	if num > ini_count+1 then
		msgbox "ini_add_comment : num が大きすぎる"
		num = ini_count + 1
	elseif num<=0 then
		num = 1
	end if
	regEx.IgnoreCase = false
	regEx.Global = False
	regEx.Pattern = "^\[.*\]$"
	if (arg_section<>"") and (not regEx.Test(arg_section)) then
		msgbox "In ini_add_section : section の入力が異常 :" & arg_section
		exit sub
	end if
	ini_insert(num)
	ini_section(num) = arg_section
	ini_key(num) = arg_section
	ini_value(num) = ""
	ini_buff(num) = arg_section
	ini_change = true
	set regEx = nothing
end sub

sub ini_insert(num)
	if UBound(ini_buff) < ini_count+1 then
		redim_ini
	end if
	dim i
	for i = ini_count to num step -1
		ini_section(i+1) = ini_section(i)
		ini_key(i+1) = ini_key(i)
		ini_value(i+1) = ini_value(i)
		ini_buff(i+1) = ini_buff(i)
	next
	ini_count = ini_count + 1
end sub

sub unpack_flag(str,flag,ndim)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。
	dim f(100),nf
	dim i,m
	for i=1 to ndim
		flag(i)=false
	next
	regEx.IgnoreCase = True
	regEx.Global = False
	regEx.Pattern = "^[ 	]*$"
	if regEx.Test(str) then
		exit sub
	end if
	split_comma str,f,nf
	for i=1 to nf
		on error resume next
		m = CInt(f(i))
		if Err.Number <> 0 then
			MsgBox "index が異常な値です : " & f(i)
	       	m=0
		end if
		On Error Goto 0
		if m >= 1 and m<=ndim then
			flag(m)=true
		elseif m = 0 then
		else
			'msgbox "index の範囲が dimension を超えています : " & m
		end if
	next
	set regEx = nothing
end sub

sub unpack_flag_2d(str,flag,m,ndim)
	dim f()
	redim f(ndim)
	dim i
	unpack_flag str,f,ndim
	for i=1 to ndim
		flag(m,i)=f(i)
	next
end sub

function pack_flag(flag,ndim)
	dim i,c,ret
	ret=""
	c=""
	for i=1 to ndim
		if flag(i) then
			ret = ret & c & CStr(i)
			c = ","
		end if
	next
	pack_flag = ret
end function

function pack_flag_2d(flag,m,ndim)
	dim f()
	redim f(ndim)
	dim i
	for i=1 to ndim
		f(i)=flag(m,i)
	next
	pack_flag_2d = pack_flag(f,ndim)
end function

sub split_comma(str,f,nf)
	dim s
	nf = 0
	s=str
	sub_split_comma s,f,nf
end sub

sub sub_split_comma(s,v,c)
	dim regEx : set regEx = New RegExp   ' 正規表現を作成します。

	'msgbox "sub_split_comma" & vbNewLine & s

	regEx.IgnoreCase = True
	regEx.Global = False
	regEx.Pattern = "^[ 	]*"
	s=regEx.Replace(s,"")

	regEx.Pattern = "^[^""^,]*[ 	]*$"
	if regEx.Test(s) then
		c = c + 1
		v(c) = s
		regEx.Pattern = "[ 	]*$"
		v(c) = regEx.Replace(v(c),"")
		exit sub
	end if

	regEx.Pattern = "^"""
	regEx.Global = False
	if regEx.Test(s) then
		c = c + 1
		s = regEx.Replace(s,"")
		regEx.Pattern = """.*$"
		v(c) = regEx.Replace(s,"")
		regEx.Pattern = "^[^""]*"""
		s = regEx.Replace(s,"")

		regEx.Pattern = "^[ 	]*,[ 	]*"
		if regEx.Test(s) then
			s = regEx.Replace(s,"")
			sub_split_comma s,v,c
		else
			regEx.Pattern = "^[^ ^	^,][^ ^	^,]*"
			if regEx.Test(s) then
				msgbox "予期しない文字列 in split_comma : " & s
			end if
		end if
		
		exit sub
	end if

	regEx.Pattern = "^[^ ^	^,]*[ 	]*,[ 	]*"
	if regEx.Test(s) then
		dim a
		regEx.Pattern = "[ 	]*,.*"
		a = regEx.Replace(s,"")
		sub_split_comma a,v,c
		regEx.Pattern = "^[^ ^	^,]*[ 	]*,[ 	]*"
		s = regEx.Replace(s,"")
		sub_split_comma s,v,c
		exit sub
	end if

	msgbox "予期しない文字列 in split_comma : " & s

	set regEx = nothing
end sub
