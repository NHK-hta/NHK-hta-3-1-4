option explicit

sub set_user_name(user,computer)
	Dim objNetWork
	'ネットワークオブジェクトの作成
	Set objNetWork = CreateObject("WScript.Network")
	'ユーザ名
	user = objNetWork.UserName
	'コンピュータ名
	computer = objNetWork.ComputerName
	Set objNetWork = Nothing
end sub

function get_file_size(filespec)
	dim fso : set fso = CreateObject("Scripting.FileSystemObject")
	Dim f
	If (fso.FileExists(filespec)) Then
		set f = fso.GetFile(filespec)
		get_file_size = CStr(f.size)
	Else
		get_file_size = "無"
	End If
	set fso = nothing
end function
