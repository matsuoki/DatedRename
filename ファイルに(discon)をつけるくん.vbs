REM 
REM �S��� �t�@�C����(discon)��t���邭�� v2.1
REM 


dim fso
dim d
dim ds
dim fpath
dim message
dim app_name

app_name="�t�@�C����(discon)�����邭��_�S���"

set fso=wscript.createobject("scripting.FileSystemObject")


if Wscript.Arguments.count=0 then
	message= "�t�@�C����t�H���_�����̃X�N���v�g�Ƀh���b�O���h���b�v���Ă��������B" & VBCRLF
	message= message & "���O�̐擪��(discon)��t���ă��l�[�����܂��B" & VBCRLF
	message= message & "" & VBCRLF
	message= message & "��F" & VBCRLF
	message= message & "�@�ŐV�t�@�C��(2).xls" & VBCRLF
	message= message & "�@������" & VBCRLF
	message= message & "�@(discon)�ŐV�t�@�C��(2).xls" & VBCRLF
	message= message & "" & VBCRLF

	msgbox message ,64,app_name

end if

for i=0 to Wscript.Arguments.count -1

	fpath=Wscript.Arguments(i)

	if fso.FolderExists( fpath  ) then 
		'�t�H���_�̏ꍇ
		Set fl = fso.GetFolder( fpath )
		ds = "(discon)"
		
		fl.Name = ds & "" & fl.Name


        elseif  fso.FileExists( fpath ) then 
		'�t�@�C���̏ꍇ
		set f=FSO.GetFile( fpath )
		'd=fl.DateCreated
		d=f.DateLastModified
		ds = "(discon)"
		f.name=ds & "" &  f.name



	else
		
        end if

rem	msgbox f.DateCreated 
rem	msgbox f.DateLastModified

next
