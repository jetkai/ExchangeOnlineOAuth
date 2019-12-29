# ExchangeOnlineOAuth
PowerShell Script - Used for connecting to Exchange Online, using Modern Authentication

[Administrator PC rights should not be required to run this script - Exchange Online admin rights is required]
This script allows any Exchange Online Admin to use the Exchange Online PowerShell cmdlet, with Multi-Factor Authentication enabled. I wrote this script as an alternative to Microsoft's module that you have to install from the Exchange Admin Centre on an Internet Explorer/Edge browser. You can use this script to integreate with your own scripts, so that you can use accounts with Multi-Factor Authentication. ~ I will also be including Microsoft's own Exchange Online Module - But with an easier way of Downloading & Installing without using the IE browser.


TODO List for the Future:

~ Do not prompt for Sign-in again [Remember me] | Save AccessToken as JSON (User Option) - Read from PC & insert into script
~ If AccessToken expires, use RefreshToken to aquire a new AccessToken
~ Optimize a little more


What this script does?
1.	Installs the Microsoft ADAL PowerShell Module
2.	Prompts the user to sign-in with their Office 365 User Credentials & will prompt for Mobile/App Authetnication (if user has this)
3.	Gathers the AccessToken from the sign-in prompt and attempts to connect to Exchange Online using BasicAuth to OAuth Conversion
4.	Imports the PowerShell session for the user to use the Exchange Online cmdlet

How to run this script?
•	You can right-click the .ps1 file and click "Run with PowerShell" - This will open a new PowerShell window in NoExit mode to allow you to connect
•	You can Import the PowerShell script within an existing PowerShell window by running the example command Import-Module -Name "C:\Users\YourName\Downloads\ExchangeOAuth.ps1
•	You can also CD to the directory by running the example command cd "C:\Users\YourName\Downloads" & Run the script with the command .\ExchangeOAuth.ps1

I am getting errors!?! How do I fix them?
•	If you are receiving the message "....ExchangeOAuth.ps1 cannot be loaded because running scripts is disabled on this system..." - You'll need to run this command and you'll then be able to run the script Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
•	If you are still not able to run the script - You'll need to also run the command Install-Module -Name Microsoft.ADAL.PowerShell


[ALTERNATIVE METHOD] Microsoft's Exchange Online Module (With a TWIST):

Description: This allows you to install /or launch Microsoft's Exchange Online Module without having to navigate to https://outlook.office.com/ecp in an Internet Explorer/Edge browser & manually download it. Inspected element to gather the link and stuck this into a quick PowerShell command to download the module immediately, useful for multiple machines.

PowerShell Script:

(New-Object -ComObject "InternetExplorer.Application").navigate2("https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application")
#------Connect to Exchange Online after Module is Installed by user-------#
Connect-EXOPSSession
