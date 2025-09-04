
	Dim sendtext
	Dim objShell

	Set objShell = CreateObject("WScript.Shell")

do 

	sendtext=inputbox("入力するアレをなにして", "パスワードいれーる" ,sendtext )

	if len(sendtext)=0 then exit do
	objShell.SendKeys  "%{TAB}" 
	WScript.Sleep 500

	objShell.SendKeys  sendtext
	WScript.Sleep 500
	
loop


