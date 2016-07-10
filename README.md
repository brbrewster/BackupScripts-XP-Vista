# BackupScripts-XP-Vista
This is a set of backup scripts that are intended to back files on a Windows XP or Vista computer to a USB drive, a mapped network drive, or a Windows SMB (or Samba) network folder, depending on the script.

I haven't tried them on newer versions of Windows, so I don't know if they will work on versions of Windows that are newer than Vista.


---


<strong>FullBackup.bat</strong> copies files from a user's Desktop, Documents/My Documents folder, Firefox profile data, Internet Explorer Favorites, and Outlook data files (PST files).
 
To make it easy to backup to any logical drive, the user is prompted to enter the letter of the logical drive that files are being backed up to.

If Outlook is open when this is run, the scripts try to close Outlook gently, and waits for Outlook to close, because the PST file cannot be copied when it is open.

Instead of simply overwriting the backup PST it attempts to keep a rotating set of up to 3 previous versions of the PST file. 

FullBackup.bat overwrites existing files each time it is used, so it should not be used for recurring backups. If recurring backups are needed, use IncrementalBackup.bat

<strong>IncrementalBackup.bat</strong> does basically the same thing that FullBackup.bat does, except it only backs up files that have changed since the last time the backup scripts have run.

It is also intended to be run as a scheduled task, so the backup drive letter is set within the script, rather than prompting the user for backup drive letter.

<strong>PSTBackup.bat</strong> is functionally the same as the PST backup portion of the other two scripts, except that it relies on a named Windows network or Samba share instead of a drive letter, because I was running into some hiccups with mapped drives at the time.

This is also intended to be run as a scheduled task.
