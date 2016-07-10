cls
@ECHO OFF
::Dependencies
::closeOutlook.vbs is used to close down Outlook cleanly to prevent PST corruption
::closeOutlook.bat is used to verify that the Outlook.exe process has ended before
::trying to copy the files in the PST directory.

::This script copies files from a user's Desktop, Documents/My Documents folder, Firefox 
::profile data, Internet Explorer Favorites, and Outlook data files (PST files) 
::to a mapped network drive.
 
::To make it easy to backup to any logical drive, the user is prompted to enter the 
::letter of the logical drive that files are being backed up to.

::This script overwrites existing files each time it is used, so it should not be
::used for recurring backups. If recurring backups are needed, use IncrementalBackup.bat

::XCOPY Switches Used:
::/R 	Overwrites read-only files.
::/S	Copies directories and subdirectories except empty ones.	
::/Z 	Copies networked files in restartable mode.
::/Y	Suppresses prompting to confirm you want to overwrite an existing destination file.

::CloseOutlook.vbs tries to close Outlook Cleanly If The User Has Not Already Done So.
::The Outlook.exe process may take some time to stop after the Outlook window has been closed, 
::and open PST files cannot be copied until after the process has stopped.
::Killing the Outlook.exe process instead of closing Outlook cleanly can corrupt open PST files

START .\Library\closeOutlook.vbs

ECHO.

::Variables

set backupCommand=XCOPY /R /S /Z /Y 
set /P driveLetter=Enter the letter of the drive you are backing files up to: 


::Tries to Validate User Input
cls
COLOR 0B
ECHO.
ECHO You entered drive letter %driveletter%. If this is not the correct location
ECHO press Ctrl+C to terminate the batch job.
ECHO. 
ECHO If this is the correct location, press any key to continue.  If you would like 
ECHO to continue, but are looking for "The ANY key," and cannot find it, press Enter.
ECHO.

@pause


::Tries to make sure programs are closed before the script runs.
cls
COLOR 0C
ECHO.
ECHO Close all open programs now!!!
ECHO Files that are being accessed by open programs cannot be backed up. 
ECHO.
ECHO Export your Firefox bookmarks to the Documents or My Documents folder,
ECHO if you have not already done so.
ECHO.
ECHO For instructions on how to export your Firefox boomarks, see Firefox Help,
ECHO or visit http://support.mozilla.com/en-US/kb/Exporting+bookmarks+to+an+HTML+file

@Pause

CLS
COLOR 07


::#############################################################
::This Section Backs Up The Desktop Contents

ECHO.
ECHO ###Backing Up Desktop Contents

if not exist "%driveLetter%:\Backup\Desktop" (
	mkdir "%driveLetter%:\Backup\Desktop"
)


if exist "%userprofile%\Desktop" (
	%backupCommand% "%userprofile%\Desktop" "%driveLetter%:\Backup\Desktop"
)

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
CALL .\Library\closeOutlook.bat 
ECHO ###Backing Up Outlook Data...

::Rotates PST file to make sure 3 backup copies are kept
CALL Library\RotatePST.bat

::Checks for Outlook data backup directory, and creates it if it doesn't already exist.


if not exist "%driveLetter%:\Backup\Outlook\PST" (
	mkdir "%driveLetter%:\Backup\Outlook\PST"
)

::Checks for PST directory path in Vista, and copies PST data if found.

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

if exist "%userprofile%\AppData\Roaming\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\AppData\Roaming\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\Autocomplete Data"
)

if exist "%userprofile%\Application Data\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Application Data\Microsoft\Outlook" "%driveLetter%:\Backup\Outlook\Autocomplete Data"
) 

ECHO.
ECHO Backup Complete
@pause
