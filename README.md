# MEMCM-SCCM-Client-Repair-and-Windows-Update-Reset
# DESCRIPTION:
PowerShell Script to uninstall and reinstall the MEMCM\SCCM Client and reset Windows update

# SYNOPSIS:
After specifying the location for MEMCM\SCCM client install files. This script will prompt the user to enter a selection to run against a single computer/Comma seperated list, or run against a list of computer names from a text file.
User will be prompted to enter credentials that have administrator rights on the computers specified.

The script will: 
  - Copy the client install files locally to each computer specified
  - Run an uninstall of the client 

  Stop and Restart each of the following services:
  - wuauserv
  - CryptSvc
  - bits
  - msiserver

  Rename the following folders:
  - "C:\Windows\SoftwareDistribution" to "C:\Windows\SoftwareDistributionBackup"
  - "C:\Windows\System32\catroot2" to "C:\Windows\System32\catroot2Backup"
  
  - Reinstall the MEMCM\SCCM Client
  - Remove the Client isntall files from the computer
  - Write an ouptput file with the results from each computer


