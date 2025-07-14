# Sysadmin-datascript
This is a powershell script that works in powershell 5.1, which is installed in every windows 10 pc.  So this script runs on every supported windows computer.
It's intention is to be used in a domain with Active Directory.  It gets all enabled computers from Active Directory and runs then a whole bunch of powershell commands remotely to them, to gather all kinds of data. Think about Get-Computerinfo, Get-Network and custom commands to see the number of temp files, both in c:/windows and in the userprofile.
This data is saved in subfolders of the c:/ (they will be made if they don't exist) in separate textfiles per computer.  The most interesting data will then be copied into an excel file.
Not everything works yet, but most does.
The idea is to have a simple overview of computers in a domain without needing to pay for hackable software.
