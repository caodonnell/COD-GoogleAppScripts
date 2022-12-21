# Google Drive Audit scripts
These functions are designed to audit the files and folders within a Google Drive folder/directory. **A big caveat is that Google Apps Scripts 
may have a short timeout period, which can limit the functionality of these scripts.** If you are in a Google Workspace account (e.g., one
that is owned by an organization, univeristy, etc. so your email doesn't end in ```@gmail.com```), scripts will timeout after 30 minutes.
However, a personal account may timeout after only 6 minutes, so if you have many files, subfiles, and/or subfolders, the script may not be
able to complete auditing your folder. The only workaround I've come up with so far is to audit smaller "chunks" (e.g., instead of doing a 
whole folder, audit the individual subfolders, or even go one level deeper). **A second caveat is that this script will only audit files/folders
to which you have at least "View" access.** If there are files/folders within that you don't have access to, they will be ommitted from the output.

