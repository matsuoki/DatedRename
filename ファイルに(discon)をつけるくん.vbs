REM 
REM ４代目 ファイルに(discon)を付けるくん v2.1
REM 


dim fso
dim d
dim ds
dim fpath
dim message
dim app_name

app_name="ファイルに(discon)をつけるくん_４代目"

set fso=wscript.createobject("scripting.FileSystemObject")


if Wscript.Arguments.count=0 then
	message= "ファイルやフォルダをこのスクリプトにドラッグ＆ドロップしてください。" & VBCRLF
	message= message & "名前の先頭に(discon)を付けてリネームします。" & VBCRLF
	message= message & "" & VBCRLF
	message= message & "例：" & VBCRLF
	message= message & "　最新ファイル(2).xls" & VBCRLF
	message= message & "　↓↓↓" & VBCRLF
	message= message & "　(discon)最新ファイル(2).xls" & VBCRLF
	message= message & "" & VBCRLF

	msgbox message ,64,app_name

end if

for i=0 to Wscript.Arguments.count -1

	fpath=Wscript.Arguments(i)

	if fso.FolderExists( fpath  ) then 
		'フォルダの場合
		Set fl = fso.GetFolder( fpath )
		ds = "(discon)"
		
		fl.Name = ds & "" & fl.Name


        elseif  fso.FileExists( fpath ) then 
		'ファイルの場合
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
