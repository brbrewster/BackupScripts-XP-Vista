@ECHO OFF
::This script performs the same function as RotatePST.bat. The reason this was created is that there are 
::instances where the backup path will be different from the root directory of a local or network drive. In 
::these cases the colon after the drive letter variable in RotatePST.bat prevents the script from functioning.
::
::In the place of the %driveletter% variable, this file relies on the %backuppath% variable. The %backuppath%
::variable should be created and assigned a value before RotatePST_nodriveletter.bat is called.


if exist "%backuppath%\Backup\Outlook\PST\Outlook_Backup_2.pst" (
	if exist "%backuppath%\Backup\Outlook\PST\Outlook_Backup_3.pst" (
		del "%backuppath%\Backup\Outlook\PST\Outlook_Backup_3.pst"		
	)

	ren "%backuppath%\Backup\Outlook\PST\Outlook_Backup_2.pst" Outlook_Backup_3.pst	 
)


if exist "%backuppath%\Backup\Outlook\PST\Outlook.pst" (
	ren "%backuppath%\Backup\Outlook\PST\Outlook.pst" Outlook_Backup_2.pst
)