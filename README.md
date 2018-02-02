# Powershell-Quickbooks
Connect to Quickbooks directly from Powershell (x86)

## Instructions
Quickbooks SDK needed for installation of QBFC13 and/or QBXMLRP2  
Open Powershell and navigate to folder that has this script. then `.\powershell_qbfc.ps1`  
If you're running in 64-bit it should relaunch as 32-bit


### Potential Problems
If you try to run as a script and get an Execution Policy error, run this:
```powershell
Set-ExecutionPolicy RemoteSigned
```
Also if you've never connected before you will need to be logged in as Admin in quickbooks and allow the script to connect.  
