AD-Locked-Out-Users
===================

Powershell Script for Help Desk to show currently locked out users

This is a simple powershell script that displays all currently locked out users in the current domain and it allow you to just double click to unlock the accounts. This program is very useful for a help desk employee because it can quickly speed up the unending calls to unlock accounts. It can also be useful for an Administrators to view which accounts are being bruteforced.

For Windows 7 systems you will want to apply the windows hotfix to your system. This will allow you to unlock accounts. If you are unable to unlock accounts chances are you haven't applied the hotfix.

You will also have to have Active Directory Tools enabled on your system for this script to work. If you are able to unlock accounts currently, then you probably already have this enabled.

The executable is just the powershell script compiled with PS2EXE. When compiling the program I recommend you use a "-noconsole" parameter.

If you have any issues feel free to email me at CalebCoffie@gmail.com or through my website at https://www.CalebCoffie.com
