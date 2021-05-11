REM 
REM ４代目 ファイルに日付を付けるくん v2
REM 


dim fso
dim d
dim ds
dim fpath
set fso=wscript.createobject("scripting.FileSystemObject")

for i=0 to Wscript.Arguments.count -1

	fpath=Wscript.Arguments(i)

	if fso.FolderExists( fpath  ) then 
		'フォルダの場合
		Set fl = fso.GetFolder( fpath )
		d=fl.DateCreated
		ds = year(d) &  right("0" &  month(d),2) & right("0" & day(d) ,2)
		fl.Name = ds & "_" & fl.Name

        elseif  fso.FileExists( fpath ) then 
		'ファイルの場合

		set f=FSO.GetFile( fpath )
		d=f.DateCreated
		ds = year(d) &  right("0" &  month(d),2) & right("0" & day(d) ,2)
		f.name=ds & "_" &  f.name

	else
		
        end if

rem	msgbox f.DateCreated 
rem	msgbox f.DateLastModified

next
