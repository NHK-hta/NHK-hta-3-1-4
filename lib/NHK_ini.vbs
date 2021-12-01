option explicit

const ini_file="NHK.ini"
const section_general="[general]"
const key_user="user"
const key_computer="computer"
const key_unicode="unicode"
'const key_delete="delete"
'const key_flv="flv_folder"
const key_mp3="mp3_folder"
const key_sub_folder="mp3_sub_folder"
const key_m4a="m4a"
const key_aac="aac"
const key_kouza="kouza"
const key_check="check"
const section_random="[random]"
const key_random_code="code"
const key_random_date="date"
const key_original_name="original_name"
const key_mp3den="mp3_density"

sub option_initialize
	dim ini_user,ini_computer,num,str
	if read_ini_file = 0 then
		ini_add "",";",";	NHK.hta  初期値設定-file"
		ini_add section_general, key_user, user_name
		ini_add section_general, key_computer, computer_name
		option_write
		ini_add "[ボキャブライダー]", "check", "1,2,3,4,5"
	else
		ini_search section_general,key_user,ini_user,num
		ini_search section_general,key_computer,ini_computer,num
		if (ini_user <> user_name) or (ini_computer <> computer_name) then
			msgbox "ユーザー名:" & ini_user & "->" & user_name & vbNewLine & _
					"コンピュータ名" & ini_computer & "->" & computer_name & vbNewLine & _
					"mp3のフォルダを初期値に戻します"
					'	"mp3とflvのフォルダを初期値に戻します"
			ini_add section_general, key_user, user_name
			ini_add section_general, key_computer, computer_name
			'ini_add section_general, key_flv, """" & flv_default & """"
			ini_add section_general, key_mp3, """" & mp3_default & """"
		end if
	end if
	option_read
	kouza_read
end sub


sub option_write
	ini_add section_general, key_unicode, CStr(unicode_flag)
	'ini_add section_general, key_delete, CStr(flv_delete_flag)
	'ini_add section_general, key_flv, """" & flv_folder & """"
	ini_add section_general, key_mp3, """" & mp3_folder & """"
	ini_add section_general, key_sub_folder, CStr(sub_folder_flag)
	ini_add section_general, key_m4a, CStr(m4a_flag)
	ini_add section_general, key_aac, CStr(aac_flag)
	ini_add section_general, key_original_name, CStr(original_name)
	ini_add section_general, key_mp3den, CStr(mp3den)
end sub

sub kouza_write
	ini_add section_general, key_kouza, pack_flag(kouza_select,kouza_dim)
end sub

sub option_read
	flag_read section_general, key_unicode, "unicode", unicode_flag, false
	'flag_read section_general, key_delete, "flv-delete", flv_delete_flag, false
	'str_read section_general, key_flv, "flv-folder", flv_folder, flv_default
	str_read section_general, key_mp3, "mp3-folder", mp3_folder, mp3_default
	flag_read section_general, key_sub_folder, "sub-folder", sub_folder_flag, false
	flag_read section_general, key_m4a, "m4a", m4a_flag, false
	flag_read section_general, key_aac, "aac", aac_flag, false
	flag_read section_general, key_original_name, "original-name", original_name, false
	str_read section_general, key_mp3den, "mp3-density", mp3den, mp3high
end sub

sub kouza_read
	dim	str,num
	ini_search section_general, key_kouza, str, num
	unpack_flag str, kouza_select, nkouza
end sub



sub ini_write_random(code,set_date)
	ini_add section_random, key_random_code, """" & code & """"
	ini_add section_random, key_random_date, """" & set_date & """"
end sub

sub ini_read_random(code,set_date)
	str_read section_random, key_random_code, "random code", code, ""
	str_read section_random, key_random_date, "date of random code", set_date, ""
end sub
