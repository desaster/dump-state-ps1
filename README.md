# dump-state.ps1

This powershell gathers various bits of information on a Windows computer and
dumps them to csv files in a timestamped folder.

The purpose of this is to help keeping a journal of installed stuff for
keeping things clean.

Things that are written out:

* Drivers
* Applications
* Services
* Program folders
* Start menu folders
* Startup programs
* BIOS version

To make running the script easier, create a shortcut of the .ps1 file and:
* Change the target to "powershell.exe C:\path-to-script\dump-state.ps1", 
* Click Advanced and select "Run As administrator"
