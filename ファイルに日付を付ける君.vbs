REM 
REM ４代目 ファイルに日付を付けるくん v2.1
REM 


dim fso
dim d
dim ds
dim fpath
dim message
dim app_name

app_name="ファイルに最終更新日をつけるくん_４代目"

set fso=wscript.createobject("scripting.FileSystemObject")


if Wscript.Arguments.count=0 then
	message= "ファイルやフォルダをこのスクリプトにドラッグ＆ドロップしてください。" & VBCRLF
	message= message & "名前の先頭に作成日のYYYYMMDDを付けてリネームします。" & VBCRLF
	message= message & "" & VBCRLF
	message= message & "例：" & VBCRLF
	message= message & "　最新ファイル(2).xls" & VBCRLF
	message= message & "　↓↓↓" & VBCRLF
	message= message & "　20210401_最新ファイル(2).xls" & VBCRLF
	message= message & "" & VBCRLF

	msgbox message ,64,app_name

end if

for i=0 to Wscript.Arguments.count -1

	fpath=Wscript.Arguments(i)

	if fso.FolderExists( fpath  ) then 
		'フォルダの場合
		Set fl = fso.GetFolder( fpath )
		'd=fl.DateCreated
		d=fl.DateLastModified
		ds = year(d) &  right("0" &  month(d),2) & right("0" & day(d) ,2)
		
		if left(fl.Name,1)="_" then 
			fl.Name = ds & "" & fl.Name
		else
			fl.Name = ds & "_" & fl.Name
		end if


        elseif  fso.FileExists( fpath ) then 
		'ファイルの場合

		set f=FSO.GetFile( fpath )
		'd=fl.DateCreated
		d=f.DateLastModified
		ds = year(d) &  right("0" &  month(d),2) & right("0" & day(d) ,2)

		'f.name=ds & "_" &  f.name

		if left(f.name, 1)="_" then 
			f.name=ds & "" &  f.name
		else
			f.name=ds & "_" &  f.name
		end if



	else
		
        end if

rem	msgbox f.DateCreated 
rem	msgbox f.DateLastModified

next
