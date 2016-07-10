cls
@ECHO OFF

::Dependencies
::closeOutlook.vbs is used to close down Outlook cleanly to prevent PST corruption
::closeOutlook.bat is used to verify that the Outlook.exe process has ended before
::trying to copy the files in the PST directory.

::This script copies files in a user's Documents/My Documents folder, Firefox 
::profile data, Internet Explorer Favorites, and Outlook data files (PST files) 
::to a mapped network drive.
 
::To make it easy to backup to any logical drive, the drive letter is stored in
::the %driveLetter% variable. The default drive letter is P, because this is 
::the letter that has been used to map the PST directory for most users. 

::To use this to backup to something other than the users U home directory, 
::change the letter set to be stored in the %driveLetter% variable to the letter
::of the drive you are copying to.

::XCOPY Switches Used:
::/M  Copies only files with the archive attribute set, turns off the archive attribute.
::/R 	Overwrites read-only files.
::/S	Copies directories and subdirectories except empty ones.	
::/Z 	Copies networked files in restartable mode.
::/D 	Copies files changed on or after the specified date. If no date is given, copies only those files whose
::source time is newer than the destination time.
::/Y	Suppresses prompting to confirm you want to overwrite an existing destination file.

::CloseOutlook.vbs tries to close Outlook Cleanly If The User Has Not Already Done So.
::The Outlook.exe process may take some time to stop after the Outlook window has been closed, 
::and open PST files cannot be copied until after the process has stopped.
::Killing the Outlook.exe process instead of closing Outlook cleanly can corrupt open PST files


START .\Library\closeOutlook.vbs

::Variables

set backupCommand=XCOPY /R /S /Z /D /M /Y 
set driveLetter=X




::#############################################################
::This section backs up the Documents\My Documents folder

ECHO.
ECHO ###Backing Up and Verifying Files...


if not exist "%driveLetter%:\Backup\Documents" (
	mkdir "%driveLetter%:\Backup\Documents"
)

::Checks for and copies files in Vista Documents folder
if exist "%userprofile%\Documents" (
	%backupCommand% "%userprofile%\Documents" "%driveLetter%:\Backup\Documents"
)


::Checks for and copies files XP Documents
if exist "%userprofile%\My Documents" (
	%backupCommand% "%userprofile%\My Documents" "%driveLetter%:\Backup\Documents"
)



::#############################################################
::This section backs up Firefox data

ECHO ###Backing Up Firefox Profile

if not exist "%driveLetter%:\Backup\Firefox" (
	mkdir "%driveLetter%:\Backup\Firefox"
)

if exist "%userprofile%\AppData\Roaming\Mozilla\Firefox\Profiles" (
	%backupCommand% "%userprofile%\AppData\Roaming\Mozilla\Firefox\Profiles" "%driveLetter%:\Backup\Firefox"
)


if exist "%userprofile%\Application Data\Mozilla\Firefox" (
	%backupCommand% "%userprofile%\Application Data\Mozilla\Firefox" "%driveLetter%:\Backup\Firefox"
)


::###############################################################
::This section backs up Internet Explorer Favorites

ECHO ###Backing up Internet Exploder Favorites

if not exist "%driveLetter%:\Backup\IE Favorites" (
	mkdir "%driveLetter%:\Backup\IE Favorites"
)


if exist "%userprofile%\Favorites" (
	%backupCommand% "%userprofile%\Favorites" "%driveLetter%:\Backup\IE Favorites"	
)





::#############################################################
::This section backs up PST Info


ECHO ###Making Sure Outlook is Closed Before Backing Up PST
CALL Library\closeOutlook.bat 
ECHO ###Backing Up Outlook Data...


::Rotates PST file to make sure 3 backup copies are kept
CALL Library\RotatePST.bat

::Checks for Outlook data backup directory, and creates it if it doesn't already exist.


if not exist "%driveLetter%:\Backup\Outlook\PST" (
	mkdir "%driveLetter%:\Backup\Outlook\PST"
)

::Checks for PST directory path in Vista, and copies PST data if found. If it is not found,
::it checkes for PST directory path on XP. Vista has files in both directory locations, so 
::it is best to keep this as an if else statement.

if exist "%userprofile%\AppData\Local\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\AppData\Local\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\PST"
) else (
  	%backupCommand% "%userprofile%\Local Settings\Application Data\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\PST"
) 



::###########################################################
::This section backs Up Outlook Autocomplete Info

::Checks for Outlook autocomplete backup directory, and creates one if it does not exist.

if not exist "%driveLetter%:\Backup\Outlook\Autocomplete Data" (
	mkdir "%driveLetter%:\Backup\Outlook\Autocomplete Data"
)

::Checks for NK2 (autocomplete data) directory on Vista, and copies files if found.

if exist "%userprofile%\AppData\Roaming\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\AppData\Roaming\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\Autocomplete Data"
)

::Checks for NK2 directory on XP, and copies files if found.

if exist "%userprofile%\Application Data\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Application Data\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\Autocomplete Data"
) 

