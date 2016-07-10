:CheckOutlook
start .\Library\closeOutlook.vbs
tasklist > tasklist.txt
for /f "tokens=1 delims= " %%i in (tasklist.txt) do (
	if %%i == OUTLOOK.EXE ( 
		ping -n 12 127.0.0.1
		goto CheckOutlook)

)
del tasklist.txt