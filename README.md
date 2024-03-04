# CoreFixes
 checks for issues in core tools we use and attempts to fix them
# common solutions for first runs
If it fails saying scripts can't be run:
set-executionpolicy remotesigned 
then Y when it asks how to change it.
If it is just sitting on the winget update task press y and enter.  It is acutally asking if you agree to the souce agreement terms. if you want to actually see the original prompt open a powershell window and do winget list instead
# note you can not run this in the blue powershell console as admin, it will open multiple windows.
