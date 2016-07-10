cls
@ECHO ON
::Dependencies
::closeOutlook.vbs is used to close down Outlook cleanly to prevent PST corruption
::closeOutlook.bat is used to verify that the Outlook.exe process has ended before
::trying to copy the files in the PST directory.

::This script copies and Outlook data files (PST files) 
::to a Samba share on a windows network.
 

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

SET backupCommand=XCOPY /R /S /Z /Y 

::Sets the name of the Samba Share in a variable instead of the mapped drive letter 
::in case there are problems with the drive mapping.
SET filesrv=ReplaceThisTextWithTheNameOfYourSambaShare

::#############################################################
::This section backs up PST Info

ECHO ###Making Sure Outlook is Closed Before Backing Up PST
CALL Library\closeOutlook.bat 

::Rotates PST file to make sure 3 backup copies are kept
::RotatePST_nodriveletter.bat relies on the %backuppath% variable to supply the path to the backup directory
SET backuppath="\\%filesrv%\PST\Home\%userid%"
CALL Library\RotatePST_nodriveletter.bat


::Checks for Outlook data backup directory, and creates it if it doesn't already exist.
if not exist "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\PST\" (
	mkdir "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\PST\"
)

::Checks for PST directory path in Vista, and copies PST data if found.



if exist "%userprofile%\AppData\Local\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\AppData\Local\Microsoft\Outlook" "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\PST\"
) else (
  	%backupCommand% "%userprofile%\Local Settings\Application Data\Microsoft\Outlook" "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\PST\"
) 



::###########################################################
::This section backs Up Outlook Autocomplete Info

::Checks for Outlook autocomplete backup directory, and creates one if it does not exist.

if not exist "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\Autocomplete Data\" (
	mkdir "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\Autocomplete Data\"
)

if exist "%userprofile%\AppData\Roaming\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\AppData\Roaming\Microsoft\Outlook" "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\Autocomplete Data\"
)

if exist "%userprofile%\Application Data\Microsoft\Outlook" (
	%backupCommand% "%userprofile%\Application Data\Microsoft\Outlook" "\\%filesrv%\PST\Home\%userid%\Backup\Outlook\Autocomplete Data\"
) 

ECHO.
ECHO Backup Complete
@pause
