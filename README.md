# CoreFixes
 checks for issues in core tools we use and attempts to fix them
# Script here
[Downlad here](https://github.com/mrdatawolf/CoreFixes/raw/main/coreFixes.ps1)
## Common solutions for first runs
### If it fails saying scripts can't be run:
set-executionpolicy remotesigned 
then Y when it asks how to change it.
### If it is just sitting on the winget update task 
press y and enter.  It is actually asking if you agree to the souce agreement terms. if you want to actually see the original prompt open a powershell window and do winget list instead.
### If it closes right away or you see a ExecutionPolicy error
1. Win10 Pro
* Set-ExecutionPolicy Unrestricted
2. Win11 Pro
* Set-ExecutionPolicy -Scope CurrentUser Unrestricted

## You can also try
From (https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing?view=powershell-7.4):
To run an unsigned script, use the Unblock-File cmdlet or use the following procedure.
1. Save the script file on your computer.
2. Click Start, click My Computer, and locate the saved script file.
3. Right-click the script file, and then click Properties.
4. Click Unblock.
# note you can not run this in the blue powershell console as admin, it will open multiple windows.
<!-- INSTALL_COMMAND: curl -L -o coreFixes.ps1 https://github.com/mrdatawolf/CoreFixes/raw/main/coreFixes.ps1 -->
<!-- RUN_COMMAND: ./coreFixes.ps1 -->
