option explicit

sub set_user_name(user,computer)
	Dim objNetWork
	'�l�b�g���[�N�I�u�W�F�N�g�̍쐬
	Set objNetWork = CreateObject("WScript.Network")
	'���[�U��
	user = objNetWork.UserName
	'�R���s���[�^��
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
		get_file_size = "��"
	End If
	set fso = nothing
end function
