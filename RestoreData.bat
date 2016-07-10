@ECHO OFF
::This script copies files from one directory or logical drive to another. 
::If it is being used to copy files from one drive to another, it can be 
::used for incremental backups without editing the %backupCommand% variable.
::If it is being used to backup files one time %backupCommand% variable should 
::not need to be changed.

::If it is being used to do recurring incremental backups to a Samba share,
::the /M switch will need to be added to xcopy in the %backupCommand% variable

::To make it easy to backup to any logical drive, the drive letter is stored in
::the %driveLetter% variable. The default drive letter is U, because this is 
::the letter that as been used to map the home directory for most users. To use
::this to backup to something other than the users U home directory, simply 
::change the letter set to be stored in the %driveLetter% variable to the one
::we are copying to.

::XCOPY Switches Used:
::/B  Copies the Symbolic Link itself versus the target of the link.
::/R 	Overwrites read-only files.
::/S	Copies directories and subdirectories except empty ones.	
::/Z 	Copies networked files in restartable mode.
::/Y	Suppresses prompting to confirm you want to overwrite an existing destination file.


set backupCommand=XCOPY /B /R /S /Z /Y 
set /P driveLetter=Enter the backup drive letter: 

::#############################################################
::This section copies the backedup files to the desktop

ECHO.
ECHO ###Copying Backup Folder to the Desktop

if not exist "%userprofile%\Desktop\Backup" (
	mkdir "%userprofile%\Desktop\Backup" 
)
%backupCommand% "%driveLetter%:\Backup" "%userprofile%\Desktop\Backup"


::#############################################################
::This section restores files to the Documents\My Documents folder

ECHO.
ECHO ###Restoring The Documents\My Documents Contents

if exist "%userprofile%\Documents" (
	%backupCommand% "%userprofile%\Desktop\Backup\Documents" "%userprofile%\Documents" 
)



if exist "%userprofile%\My Documents" (
	%backupCommand% "%userprofile%\Desktop\Backup\Documents" "%userprofile%\My Documents" 
)


ECHO ###Restoring Outlook Data
CALL .\Library\CloseOutlook.bat
::#############################################################
::This section Restores PST data

::Checks for PST directory path in Vista, and copies PST data if found.
::if not found, it checks for the PST directory path in XP, and copies it there

if exist "%userprofile%\AppData\Local\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Desktop\Backup\Outlook\PST" "%userprofile%\AppData\Local\Microsoft\Outlook" 
) else (
  	%backupCommand% "%userprofile%\Desktop\Backup\Outlook\PST" "%userprofile%\Local Settings\Application Data\Microsoft\Outlook" 
) 



::###########################################################
::This section restores Outlook Autocomplete Info

::Checks for Outlook autocomplete backup directory, and creates one if it does not exist.

if exist "%userprofile%\AppData\Roaming\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Desktop\Backup\Outlook\Autocomplete Data" "%userprofile%\AppData\Roaming\Microsoft\Outlook" 
)

if exist "%userprofile%\Application Data\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Desktop\Backup\Outlook\Autocomplete Data"  "%userprofile%\Application Data\Microsoft\Outlook"
) 


::###############################################################
::This section restores Internet Explorer Favorites

ECHO ###Backing up Internet Exploder Favorites

if exist "%userprofile%\Favorites" (
	%backupCommand% "%userprofile%\Desktop\Backup\IE Favorites" "%userprofile%\Favorites" 	
)

::#############################################################
ECHO.
ECHO Because of the way Firefox handles profile data, Firefox bookmarks
ECHO cannot be restored using an automated script. To restore your Firefox 
ECHO bookmarks, import them from the bookmarks.htm file if you exported your
ECHO bookmarks to back them up.
ECHO.
ECHO If you did not export your bookmarks to a bookmarks.htm file, and were 
ECHO you may be able restore them from a .JSON file in your backed up Firefox 
ECHO profile.


@pause
