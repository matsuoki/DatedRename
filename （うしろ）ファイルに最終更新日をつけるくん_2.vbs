REM 
REM �S��� �t�@�C���ɓ��t��t���邭�� v2.1
REM 


dim fso
dim d
dim ds
dim fpath
dim message
dim app_name

app_name="�t�@�C���ɍŏI�X�V�������邭��_�S���"

set fso=wscript.createobject("scripting.FileSystemObject")


if Wscript.Arguments.count=0 then
	message= "�t�@�C����t�H���_�����̃X�N���v�g�Ƀh���b�O���h���b�v���Ă��������B" & VBCRLF
	message= message & "���O�̐擪�ɍ쐬����YYYYMMDD��t���ă��l�[�����܂��B" & VBCRLF
	message= message & "" & VBCRLF
	message= message & "��F" & VBCRLF
	message= message & "�@�ŐV�t�@�C��(2).xls" & VBCRLF
	message= message & "�@������" & VBCRLF
	message= message & "�@�ŐV�t�@�C��(2)_2021-04-01_20-13-01-21.xls" & VBCRLF
	message= message & "" & VBCRLF

	msgbox message ,64,app_name

end if

for i=0 to Wscript.Arguments.count -1

	fpath=Wscript.Arguments(i)

	if fso.FolderExists( fpath ) then 

		'�t�H���_�̏ꍇ
		Set fl = fso.GetFolder( fpath )

		'd=fl.DateCreated
		d=fl.DateLastModified

		ds = year(d) & "-" & right("0" &  month(d),2) &  "-" & right("0" & day(d) ,2)  & _
			"_" & hour(d) & "-" & right("0" &  minute(d),2) &  "-" & right("0" & second(d) ,2)  
		
		if right(fl.Name,1)="_" then 
			fl.Name =  fl.Name & ds
		else
			fl.Name = fl.Name & "_" & ds
		end if


        elseif  fso.FileExists( fpath ) then 
		'�t�@�C���̏ꍇ

		set f=FSO.GetFile( fpath )
		'd=fl.DateCreated
		d=f.DateLastModified

		ds = year(d) & "-" & right("0" &  month(d),2) &  "-" & right("0" & day(d) ,2)  & _
			"_" & hour(d) & "-" & right("0" &  minute(d),2) &  "-" & right("0" & second(d) ,2)  

		bname =  fso.GetBaseName(fpath)
		ename = fso.GetExtensionName(fpath)


		if right(f.name, 1)="_" then 
			f.name= bname & ds & "." & ename
		else
			f.name= bname & "_" &  ds & "." & ename
		end if

	else

        end if

rem	msgbox f.DateCreated 
rem	msgbox f.DateLastModified

next
