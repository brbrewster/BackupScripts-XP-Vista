@echo off
::This script module creates a rotating set of backup copies of the Outlook.pst file. Even when Outlook is 
::closed properly, large PST files may not close cleanly. This creates the potential for large PST files to
::become corrupted.

::The backup script copies the existing Outlook.pst file to the backup location. If a PST file becomes
::corrupted when Outlook is closed, the corrupted file would then be copied over to the backup location, and
::the existing good backup file would be overwritten.

::The script creates 3 rotating copies of the Outlook.pst file to minimize the chance of data loss in the
::event the backup script is run more than once after the PST file has become corrupted. The long file names 
::are intended to reduce the risk of accidentally overwriting other valid backup PST files with the same name. 

::This nested if statement first checks to see if a 2nd copy of the PST exists before deleting the third copy,
::and renaming Outlook_Backup_2.pst to Outlook_Backup_3.pst. Without this check, the script can delete all 
::backup PST files if it is run multiple times, and the backup Outlook.pst file is not replaced by a new copy 

::RotatePST.bat relies on the %driveletter% variable, which should be created and assigned a value by the 
::calling script before RotatePST.bat is called. 

if exist "%driveLetter%:\Backup\Outlook\PST\Outlook_Backup_2.pst" (
	if exist "%driveLetter%:\Backup\Outlook\PST\Outlook_Backup_3.pst" (
		del "%driveLetter%:\Backup\Outlook\PST\Outlook_Backup_3.pst"		
	)

	ren "%driveLetter%:\Backup\Outlook\PST\Outlook_Backup_2.pst" Outlook_Backup_3.pst	 
)


if exist "%driveLetter%:\Backup\Outlook\PST\Outlook.pst" (
	ren "%driveLetter%:\Backup\Outlook\PST\Outlook.pst" Outlook_Backup_2.pst
)